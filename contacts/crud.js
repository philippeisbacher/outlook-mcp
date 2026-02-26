/**
 * Contacts CRUD functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

const CONTACT_SELECT = 'id,displayName,emailAddresses,jobTitle,department,mobilePhone,businessPhones,officeLocation';

function formatContact(c) {
  const email = c.emailAddresses?.[0]?.address || 'no email';
  const phone = c.mobilePhone || c.businessPhones?.[0] || '';
  const title = c.jobTitle ? `\n   ${c.jobTitle}${c.department ? ` — ${c.department}` : ''}` : '';
  const phoneStr = phone ? `\n   ${phone}` : '';
  return `${c.displayName}\n   ${email}${title}${phoneStr}`;
}

function authError() {
  return {
    content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }],
    isError: true
  };
}

// ── list-contacts ─────────────────────────────────────────────────────────────
async function handleListContacts(args) {
  const { count = 25 } = args;

  try {
    const accessToken = await ensureAuthenticated();

    const response = await callGraphAPI(accessToken, 'GET', 'me/contacts', null, {
      $top: Math.min(count, config.MAX_RESULT_COUNT),
      $select: CONTACT_SELECT,
      $orderby: 'displayName asc'
    });

    const contacts = response.value || [];

    if (contacts.length === 0) {
      return { content: [{ type: "text", text: "No contacts found." }] };
    }

    const list = contacts.map((c, i) => `${i + 1}. ${formatContact(c)}`).join('\n\n');
    return {
      content: [{ type: "text", text: `Found ${contacts.length} contact(s):\n\n${list}` }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') return authError();
    return { content: [{ type: "text", text: `Error listing contacts: ${error.message}` }], isError: true };
  }
}

// ── get-contact ───────────────────────────────────────────────────────────────
async function handleGetContact(args) {
  const { id } = args;

  if (!id) {
    return { content: [{ type: "text", text: "Contact ID is required." }], isError: true };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const contact = await callGraphAPI(accessToken, 'GET', `me/contacts/${id}`, null, {
      $select: CONTACT_SELECT
    });

    return {
      content: [{ type: "text", text: `Contact details:\n\n${formatContact(contact)}\nID: ${contact.id}` }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') return authError();
    return { content: [{ type: "text", text: `Error getting contact: ${error.message}` }], isError: true };
  }
}

// ── create-contact ────────────────────────────────────────────────────────────
async function handleCreateContact(args) {
  const { displayName, email, jobTitle, department, mobilePhone } = args;

  if (!displayName) {
    return { content: [{ type: "text", text: "displayName is required." }], isError: true };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const body = { displayName };
    if (email) body.emailAddresses = [{ address: email, name: displayName }];
    if (jobTitle) body.jobTitle = jobTitle;
    if (department) body.department = department;
    if (mobilePhone) body.mobilePhone = mobilePhone;

    const contact = await callGraphAPI(accessToken, 'POST', 'me/contacts', body);

    return {
      content: [{ type: "text", text: `Contact '${displayName}' created. ID: ${contact.id}` }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') return authError();
    return { content: [{ type: "text", text: `Error creating contact: ${error.message}` }], isError: true };
  }
}

// ── update-contact ────────────────────────────────────────────────────────────
async function handleUpdateContact(args) {
  const { id, displayName, email, jobTitle, department, mobilePhone } = args;

  if (!id) {
    return { content: [{ type: "text", text: "Contact ID is required." }], isError: true };
  }

  const patch = {};
  if (displayName !== undefined) patch.displayName = displayName;
  if (email !== undefined) patch.emailAddresses = [{ address: email, name: displayName || '' }];
  if (jobTitle !== undefined) patch.jobTitle = jobTitle;
  if (department !== undefined) patch.department = department;
  if (mobilePhone !== undefined) patch.mobilePhone = mobilePhone;

  if (Object.keys(patch).length === 0) {
    return {
      content: [{ type: "text", text: "No fields to update. Provide at least one of: displayName, email, jobTitle, department, mobilePhone." }],
      isError: true
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    await callGraphAPI(accessToken, 'PATCH', `me/contacts/${id}`, patch);

    return { content: [{ type: "text", text: `Contact ${id} updated successfully.` }] };
  } catch (error) {
    if (error.message === 'Authentication required') return authError();
    return { content: [{ type: "text", text: `Error updating contact: ${error.message}` }], isError: true };
  }
}

// ── delete-contact ────────────────────────────────────────────────────────────
async function handleDeleteContact(args) {
  const { id } = args;

  if (!id) {
    return { content: [{ type: "text", text: "Contact ID is required." }], isError: true };
  }

  try {
    const accessToken = await ensureAuthenticated();

    await callGraphAPI(accessToken, 'DELETE', `me/contacts/${id}`, null);

    return { content: [{ type: "text", text: `Contact ${id} deleted successfully.` }] };
  } catch (error) {
    if (error.message === 'Authentication required') return authError();
    return { content: [{ type: "text", text: `Error deleting contact: ${error.message}` }], isError: true };
  }
}

module.exports = {
  handleListContacts,
  handleGetContact,
  handleCreateContact,
  handleUpdateContact,
  handleDeleteContact
};
