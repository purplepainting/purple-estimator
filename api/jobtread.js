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
async function paveCall(query, grantKey) {
  const upstream = await fetch(JOBTREAD_API_URL, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ query, variables: {}, grantKey }),
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
            $: { size: 10, where: [['name', 'like', `%${query}%`]] },
            nodes: { id: {}, name: {}, primaryContact: { id: {}, name: {}, email: {}, phone: {} } },
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
            $: { size: 10, where: [['name', 'like', `%${query}%`]] },
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
    async execute(payload, grantKey) {
      const { customerId, name, address, tier, bidOrigin } = payload;
      const customFieldValues = {};
      const jt = TIER_TO_JOB_TYPE[tier];
      if (jt) customFieldValues[CUSTOM_FIELDS.JOB_TYPE] = jt;
      const bo = BID_ORIGIN_TO_LABEL[bidOrigin];
      if (bo) customFieldValues[CUSTOM_FIELDS.BID_ORIGIN] = bo;

      const args = { accountId: customerId, name };
      if (address) args.location = { address };
      if (Object.keys(customFieldValues).length > 0) args.customFieldValues = customFieldValues;

      const q = { createJob: { $: args, createdJob: { id: {}, name: {} } } };
      return paveCall(q, grantKey);
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
