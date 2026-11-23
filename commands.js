const CLIENT_ID = '6f33cd34-c1cd-45ed-8a19-3af3d765c401';
const TENANT_ID = 'c5953050-74d7-4220-91ca-8a87b24b3e72';
const REDIRECT_URI = 'https://dima-09.github.io/outlook-graph-signature/commands.html';

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: REDIRECT_URI,
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: false,
  },
};

const loginRequest = {
  scopes: ['User.Read'],
};

let msalInstance = null;
let isApplyingSignature = false;

Office.onReady(() => {
  msalInstance = new msal.PublicClientApplication(msalConfig);
});

async function getGraphUserProfile() {
  const accounts = msalInstance.getAllAccounts();
  let authResult;

  if (accounts.length > 0) {
    authResult = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0],
    });
  } else {
    authResult = await msalInstance.acquireTokenSilent(loginRequest);
  }

  const response = await fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,jobTitle', {
    headers: {
      Authorization: `Bearer ${authResult.accessToken}`,
    },
  });

  if (!response.ok) {
    throw new Error(`Graph API error: ${response.status}`);
  }

  return await response.json();
}

function buildSignatureHtml(user) {
  const displayNameUpper = escapeHtml((user.displayName || '').toLocaleUpperCase('en-US'));
  const jobTitle = escapeHtml(user.jobTitle || '');
  const mail = escapeHtml(user.mail || '');

  return `
    <div style="font-family: Arial, Calibri, sans-serif; font-size: 10pt; color: #333333; line-height: 1.35;">
      <div style="font-size: 10pt; color: #333333;">${displayNameUpper}</div>
      ${jobTitle ? `<div style="font-size: 10pt; color: #333333;">${jobTitle}</div>` : ''}
      <div style="font-size: 10pt; color: #333333;">BDO Qazaqstan</div>
      ${mail ? `<div style="font-size: 10pt; color: #333333;"><a href="mailto:${mail}" style="color: #e81a3b; text-decoration: underline;">${mail}</a></div>` : ''}

      <div style="margin-top: 14px; font-size: 10pt; color: #333333;">6A Gabdullin Street</div>
      <div style="font-size: 10pt; color: #333333;">A15H4E3, Almaty</div>
      <div style="font-size: 10pt; color: #333333;">KAZAKHSTAN</div>
      <div style="font-size: 10pt; color: #333333;">Tel: +7 727 331 31 34</div>
      <div style="font-size: 10pt; color: #333333;"><a href="https://www.bdoqz.com" target="_blank" rel="noopener noreferrer" style="color: #e81a3b; text-decoration: underline;">www.bdoqz.com</a></div>

      <div style="margin-top: 14px; font-size: 8pt; color: #333333; line-height: 1.35;">
        BDO Qazaqstan, a limited liability partnership, registered under the laws of the Republic of Kazakhstan, is a member of BDO International Limited, a UK company limited by guarantee, and forms part of the international BDO network of independent member firms. BDO is the brand name for the BDO network and for each of the BDO Member Firms.
      </div>
    </div>
  `.trim();
}

function escapeHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function setItemSignature(html) {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item || !item.body || typeof item.body.setSignatureAsync !== 'function') {
      reject(new Error('setSignatureAsync is not supported in this context'));
      return;
    }

    item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error ? result.error.message : 'Failed to set signature'));
      }
    });
  });
}

async function applySignatureEvent(event) {
  if (isApplyingSignature) {
    event.completed();
    return;
  }

  isApplyingSignature = true;
  try {
    if (!msalInstance) {
      msalInstance = new msal.PublicClientApplication(msalConfig);
    }
    // const user = await getGraphUserProfile();
    // const signatureHtml = buildSignatureHtml(user);
    // await setItemSignature(signatureHtml);
    await setItemSignature('<div><strong>TEST SIGNATURE</strong></div>');
  } catch (err) {
    console.error('Auto-signature error:', err);
  } finally {
    isApplyingSignature = false;
    event.completed();
  }
}

if (typeof Office !== 'undefined' && Office.actions) {
  Office.actions.associate('onNewMessageComposeHandler', applySignatureEvent);
  Office.actions.associate('onMessageRecipientsChangedHandler', applySignatureEvent);
}
