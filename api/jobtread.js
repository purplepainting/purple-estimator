// Vercel serverless proxy → JobTread Pave API.
// Mirrors api/chat.js patterns: raw body read (bypass Vercel JSON parser),
// CORS preflight, action-based dispatch, structured error responses.

const JOBTREAD_API_URL = 'https://api.jobtread.com/pave';
const ORG_ID = '22PWNY9u7qZd';

const CUSTOM_FIELDS = {
  JOB_TYPE:      '22PWsEVVW4aj',
  BID_ORIGIN:    '22PWsDkAPRYB',
  LEAD_SOURCE:   '22PWNYKhTU6S',
  PHONE:         '22PWNYKhWqBE',
  LEAD_STAGE:    '22PWsFuiWzEq',
};

const TIER_TO_JOB_TYPE = {
  standard:   'Residential - Standard Home',
  production: 'Property Management / Production',
  highend:    'Residential - Custom House',
  prevailing: 'Prevailing Wage',
};

const BID_ORIGIN_TO_LABEL = {
  'Job Walk':         'Requested Job Walk',
  'Digital Takeoff':  'Received Digital Bid Invite',
  'Partner Work Order': 'Partner Work Order',
};

// JT's "like" is a literal substring match. Searching the whole phrase
// ("purple painting co") fails when the stored name is just "purple painting".
// Tokenize, drop common business-suffix filler words, and OR the remaining
// words so any single token hit counts as a match.
const SEARCH_FILLER_WORDS = new Set([
  'co', 'company', 'inc', 'llc', 'the', 'family', 'residence', '&', 'and',
]);

function tokenizeForSearch(query) {
  const all = String(query || '').split(/\s+/).filter((w) => w.length > 0);
  const filtered = all.filter((w) => !SEARCH_FILLER_WORDS.has(w.toLowerCase()));
  return filtered.length > 0 ? filtered : all;
}

function nameLikeCondition(query) {
  const words = tokenizeForSearch(query);
  if (words.length === 1) return ['name', 'like', `%${words[0]}%`];
  return { or: words.map((w) => ['name', 'like', `%${w}%`]) };
}

export const config = {
  api: { bodyParser: false },
};

function readRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (chunk) => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf-8')));
    req.on('error', reject);
  });
}

// Single Pave call. Returns parsed JSON (or { _raw, _httpStatus } if not JSON).
//
// IMPORTANT: Pave expects the grantKey nested inside the query's "$" args object,
// NOT as a top-level sibling of `query`. Top-level grantKey is silently ignored
// — Pave returns {"organization":null} with HTTP 200 and the grant shows "Never
// Used" in JT. Merge grantKey into query.$ (preserving any existing $ keys).
async function paveCall(query, grantKey) {
  const authedQuery = { ...query, $: { grantKey, ...(query.$ || {}) } };
  const upstream = await fetch(JOBTREAD_API_URL, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ query: authedQuery }),
  });
  const text = await upstream.text();
  let data;
  try { data = JSON.parse(text); }
  catch { return { _httpStatus: upstream.status, _raw: text.slice(0, 1000) }; }
  data._httpStatus = upstream.status;
  return data;
}

// ── Action implementations ─────────────────────────────────────────────────

const ACTIONS = {
  find_customer: {
    required: ['query'],
    async execute({ query }, grantKey) {
      const q = {
        organization: {
          $: { id: ORG_ID },
          accounts: {
            $: {
              where: { and: [['type', 'customer'], nameLikeCondition(query)] },
              size: 10,
            },
            nodes: { id: {}, name: {}, type: {}, primaryContact: { id: {}, name: {} } },
          },
        },
      };
      return paveCall(q, grantKey);
    },
  },

  get_customer_jobs: {
    required: ['customerId'],
    async execute({ customerId }, grantKey) {
      const q = {
        account: {
          $: { id: customerId },
          jobs: {
            $: { size: 20 },
            nodes: { id: {}, name: {}, location: { address: {} } },
          },
        },
      };
      return paveCall(q, grantKey);
    },
  },

  find_job: {
    required: ['query'],
    async execute({ query }, grantKey) {
      const q = {
        organization: {
          $: { id: ORG_ID },
          jobs: {
            $: {
              where: { and: [nameLikeCondition(query)] },
              size: 10,
            },
            nodes: { id: {}, name: {}, location: { address: {} }, account: { id: {}, name: {} } },
          },
        },
      };
      return paveCall(q, grantKey);
    },
  },

  create_customer: {
    required: ['name'],
    async execute(payload, grantKey) {
      const { name, contactName, email, phone, address, leadSource } = payload;

      // Step 1 — createAccount
      const accountCustomFields = {};
      if (leadSource) accountCustomFields[CUSTOM_FIELDS.LEAD_SOURCE] = leadSource;
      const accountArgs = { organizationId: ORG_ID, type: 'customer', name };
      if (Object.keys(accountCustomFields).length > 0) accountArgs.customFieldValues = accountCustomFields;

      const accountResp = await paveCall(
        { createAccount: { $: accountArgs, createdAccount: { id: {}, name: {} } } },
        grantKey,
      );
      const accountId = accountResp?.createAccount?.createdAccount?.id;
      if (!accountId) {
        return { _failed_at: 'createAccount', _account_response: accountResp };
      }

      // Step 2 — createCustomerContact (linked to the new account)
      const contactCustomFields = {};
      if (phone) contactCustomFields[CUSTOM_FIELDS.PHONE] = phone;
      const contactArgs = { accountId, name: contactName || name };
      if (email) contactArgs.email = email;
      if (Object.keys(contactCustomFields).length > 0) contactArgs.customFieldValues = contactCustomFields;

      const contactResp = await paveCall(
        { createCustomerContact: { $: contactArgs, createdContact: { id: {}, name: {} } } },
        grantKey,
      );

      return {
        createAccount: accountResp.createAccount,
        createCustomerContact: contactResp.createCustomerContact ?? null,
        _contact_response_raw: contactResp,
        _note: address ? 'address not stored on customer — pass it to create_job' : undefined,
      };
    },
  },

  create_job: {
    required: ['customerId', 'name'],
    // JT's createJob has NO `accountId` / `location` args — it requires
    // `locationId`. A job attaches to a customer THROUGH a location, so we
    // createLocation first and then createJob with the returned id.
    async execute(payload, grantKey) {
      const { customerId, name, address, tier, bidOrigin } = payload;

      // Without an address, createLocation has nothing to work with. Return a
      // structured error so BuildChat can ask the user (mirrors create_customer's
      // _failed_at pattern).
      if (!address || !String(address).trim()) {
        return {
          _failed_at: 'address_required',
          _message: 'A full job address (street, city, state, zip) is required to create a job.',
        };
      }

      // Step 1 — createLocation under the customer. parseAddress defaults true,
      // so we send the full address string as one field and let JT split it.
      const locationResp = await paveCall(
        {
          createLocation: {
            $: { accountId: customerId, address },
            createdLocation: { id: {}, address: {} },
          },
        },
        grantKey,
      );
      const locationId = locationResp?.createLocation?.createdLocation?.id;
      if (!locationId) {
        return { _failed_at: 'createLocation', _location_response: locationResp };
      }

      // Step 2 — createJob attached to that location. Custom-field mapping is
      // unchanged: tier → Job Type, bidOrigin → How Did Bid Come In.
      const customFieldValues = {};
      const jt = TIER_TO_JOB_TYPE[tier];
      if (jt) customFieldValues[CUSTOM_FIELDS.JOB_TYPE] = jt;
      const bo = BID_ORIGIN_TO_LABEL[bidOrigin];
      if (bo) customFieldValues[CUSTOM_FIELDS.BID_ORIGIN] = bo;

      const jobArgs = { locationId, name };
      if (Object.keys(customFieldValues).length > 0) jobArgs.customFieldValues = customFieldValues;

      const jobResp = await paveCall(
        { createJob: { $: jobArgs, createdJob: { id: {}, name: {} } } },
        grantKey,
      );

      // Return shape preserves `createJob.createdJob.id` so the frontend
      // executeAction's ID-extraction path keeps working unchanged.
      return {
        createLocation: locationResp.createLocation,
        createJob: jobResp.createJob ?? null,
        _location_response_raw: locationResp,
      };
    },
  },

  create_cost_group: {
    required: ['jobId', 'name'],
    async execute(payload, grantKey) {
      const { jobId, name, parentCostGroupId, quantityFormula, unitId, quantity } = payload;
      const args = { jobId, name, showChildCosts: true, showChildren: true, showDescription: true };
      if (parentCostGroupId) args.parentCostGroupId = parentCostGroupId;
      if (quantityFormula) args.quantityFormula = quantityFormula;
      if (unitId) args.unitId = unitId;
      if (quantity != null) args.quantity = quantity;

      const q = { createCostGroup: { $: args, createdCostGroup: { id: {}, name: {} } } };
      return paveCall(q, grantKey);
    },
  },

  create_cost_item: {
    required: ['jobId', 'costGroupId', 'name'],
    async execute(payload, grantKey) {
      const {
        jobId, costGroupId, organizationCostItemId, name,
        costCodeId, costTypeId, unitId,
        quantityFormula, quantity, unitCost, unitPrice,
      } = payload;
      const args = { jobId, costGroupId, name, showDescription: true, showQuantity: true };
      if (organizationCostItemId) args.organizationCostItemId = organizationCostItemId;
      if (costCodeId) args.costCodeId = costCodeId;
      if (costTypeId) args.costTypeId = costTypeId;
      if (unitId) args.unitId = unitId;
      if (quantityFormula) args.quantityFormula = quantityFormula;
      if (quantity != null) args.quantity = quantity;
      if (unitCost != null) args.unitCost = unitCost;
      if (unitPrice != null) args.unitPrice = unitPrice;

      const q = { createCostItem: { $: args, createdCostItem: { id: {}, name: {} } } };
      return paveCall(q, grantKey);
    },
  },

  find_user: {
    required: ['name'],
    // Search internal team members (purplepainting.net, non-machine accounts)
    // by name fragment. Returns { name, emailAddress, phoneNumber } per match
    // so the chat can use them as fromName / fromEmailAddress / fromPhoneNumber
    // when creating a document.
    async execute({ name }, grantKey) {
      const q = {
        organization: {
          $: { id: ORG_ID },
          memberships: {
            $: { size: 100 },
            nodes: {
              user: { id: {}, name: {}, emailAddress: {}, phoneNumber: {}, isMachine: {} },
            },
          },
        },
      };
      const raw = await paveCall(q, grantKey);
      const nodes = raw?.organization?.memberships?.nodes || [];
      const needle = String(name || '').trim().toLowerCase();
      const matches = nodes
        .map((n) => n?.user)
        .filter((u) => u && !u.isMachine)
        .filter((u) => (u.emailAddress || '').toLowerCase().endsWith('@purplepainting.net'))
        .filter((u) => !needle || (u.name || '').toLowerCase().includes(needle))
        .map((u) => ({ name: u.name, emailAddress: u.emailAddress, phoneNumber: u.phoneNumber || null }));
      return { matches, _httpStatus: raw?._httpStatus ?? 200 };
    },
  },

  get_job_cost_items: {
    required: ['jobId'],
    // Return everything DocumentChat needs to build proven-shape document line
    // items in a single call: the job's location (name + address) plus each
    // leaf cost item with full pricing/coding fields. Each costItem.id is the
    // jobCostItemId passed to update_document's costItem lineItems.
    async execute({ jobId }, grantKey) {
      const q = {
        job: {
          $: { id: jobId },
          location: { name: {}, address: {} },
          costItems: {
            $: { size: 100 },
            nodes: {
              id: {},
              name: {},
              quantity: {},
              costCode: { id: {} },
              costType: { id: {} },
              unit: { id: {} },
              unitCost: {},
              unitPrice: {},
              costGroup: { id: {}, name: {} },
            },
          },
        },
      };
      return paveCall(q, grantKey);
    },
  },

  create_document: {
    required: ['jobId', 'name', 'type', 'fromName', 'toName', 'taxRate'],
    // Create a JobTread document. Required by JT: jobId, name, type, fromName,
    // toName, taxRate. `name` MUST exactly match a JT TEMPLATE name (e.g.
    // "Proposal", "Selections", "Change Order", "Interior Scope of work
    // proposal") — JT rejects custom names. Put the descriptive title in
    // `subject`. createDocument does NOT inherit the job's location — pass
    // jobLocationName / jobLocationAddress explicitly. dueDays XOR dueDate
    // — never set both.
    async execute(payload, grantKey) {
      const {
        jobId, name, type, fromName, toName, taxRate,
        description, footer, dueDays, dueDate,
        fromEmailAddress, fromPhoneNumber, toEmailAddress,
        jobLocationName, jobLocationAddress, subject, issueDate,
      } = payload;
      const args = { jobId, name, type, fromName, toName, taxRate };
      if (description) args.description = description;
      if (footer) args.footer = footer;
      // dueDays XOR dueDate — dueDate wins if both are somehow set.
      if (dueDate) args.dueDate = dueDate;
      else if (dueDays != null) args.dueDays = dueDays;
      if (fromEmailAddress) args.fromEmailAddress = fromEmailAddress;
      if (fromPhoneNumber) args.fromPhoneNumber = fromPhoneNumber;
      if (toEmailAddress) args.toEmailAddress = toEmailAddress;
      if (jobLocationName) args.jobLocationName = jobLocationName;
      if (jobLocationAddress) args.jobLocationAddress = jobLocationAddress;
      if (subject) args.subject = subject;
      if (issueDate) args.issueDate = issueDate;

      const q = { createDocument: { $: args, createdDocument: { id: {}, name: {} } } };
      return paveCall(q, grantKey);
    },
  },

  update_document: {
    required: ['id'],
    // Update an existing document. Any subset of fields may be passed; only
    // present ones are forwarded. lineItems mirror the budget via the
    // "existingCostItem" variant — { jobCostItemId, name?, quantity?,
    // unitPrice? } — referencing JOB cost items (NOT sourceCostItemId, NOT
    // organizationCostItemId).
    async execute(payload, grantKey) {
      const {
        id, name, fromName, toName, taxRate, dueDays,
        description, footer, fromEmailAddress, fromPhoneNumber, toEmailAddress,
        lineItems,
      } = payload;
      const args = { id };
      if (name) args.name = name;
      if (fromName) args.fromName = fromName;
      if (toName) args.toName = toName;
      if (taxRate != null) args.taxRate = taxRate;
      if (dueDays != null) args.dueDays = dueDays;
      if (description) args.description = description;
      if (footer) args.footer = footer;
      if (fromEmailAddress) args.fromEmailAddress = fromEmailAddress;
      if (fromPhoneNumber) args.fromPhoneNumber = fromPhoneNumber;
      if (toEmailAddress) args.toEmailAddress = toEmailAddress;
      if (Array.isArray(lineItems)) args.lineItems = lineItems;

      // updateDocument has NO `updatedDocument` read-back (the field does not
      // exist on the root mutation type, so reading it returns a GraphQL field
      // error that masquerades as an "intermittent" failure). Nest the
      // `document` root field with its own $: { id } to read the updated state
      // back. price is the live total after JT processes the line items.
      const q = {
        updateDocument: {
          $: args,
          document: { $: { id: args.id }, id: {}, name: {}, price: {} },
        },
      };
      return paveCall(q, grantKey);
    },
  },
};

// ── HTTP handler ──────────────────────────────────────────────────────────

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(204).end();

  if (req.method === 'GET' && req.query?.debug === '1') {
    const k = process.env.JOBTREAD_GRANT_KEY || '';
    return res.status(200).json({
      has_grant_key: !!k,
      key_length: k.length,
      key_prefix: k.slice(0, 6),
      org_id: ORG_ID,
      actions: Object.keys(ACTIONS),
      runtime: 'node-esm-rawbody',
      node_version: process.version,
    });
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed. Use POST.' });
  }

  const grantKey = process.env.JOBTREAD_GRANT_KEY;
  if (!grantKey) {
    return res.status(500).json({ error: 'JOBTREAD_GRANT_KEY not set on server' });
  }

  let rawBody;
  try {
    rawBody = await readRawBody(req);
  } catch (err) {
    console.error('jobtread proxy: readRawBody failed:', err.message);
    return res.status(400).json({ error: 'failed_to_read_body', message: err.message });
  }

  let body;
  try {
    body = JSON.parse(rawBody);
  } catch (err) {
    console.error('jobtread proxy: invalid JSON:', err.message);
    return res.status(400).json({ error: 'invalid_json', message: err.message });
  }

  const { action, payload } = body || {};
  console.log(
    'jobtread proxy: method=', req.method,
    'content-type=', req.headers['content-type'],
    'action=', action,
    'payload-keys=', payload ? Object.keys(payload) : null,
  );
  if (!action || typeof action !== 'string') {
    return res.status(400).json({ error: 'missing_action' });
  }
  const spec = ACTIONS[action];
  if (!spec) {
    return res.status(400).json({ error: 'unknown_action', action, available: Object.keys(ACTIONS) });
  }
  for (const field of spec.required) {
    if (payload == null || payload[field] == null || payload[field] === '') {
      return res.status(400).json({ error: 'missing_required_field', action, field });
    }
  }

  console.log('jobtread proxy:', action, 'payload keys:', Object.keys(payload || {}));

  try {
    const result = await spec.execute(payload, grantKey);
    const httpStatus = result?._httpStatus ?? 200;
    return res.status(httpStatus).json(result);
  } catch (err) {
    console.error('jobtread proxy fatal:', action, err.message);
    return res.status(502).json({ error: 'jobtread_fetch_failed', action, message: err.message });
  }
}
