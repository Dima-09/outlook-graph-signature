import { getUserProfile } from './graph-auth.js';
import { buildSignature } from './signature-template.js';

function setSignature(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSignatureAsync(
      html,
      {
        coercionType: Office.CoercionType.Html,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(result.error);
        }
      },
    );
  });
}

async function onNewMessageComposeHandler(event) {
  try {
    console.log('BDO: Event started');

    const user = await getUserProfile(false);

    if (!user.displayName || !user.displayName.trim()) {
      throw new Error('User display name is missing');
    }

    const signature = buildSignature(user);

    await setSignature(signature);

    console.log('BDO: Signature inserted successfully');
  } catch (error) {
    console.error('BDO: Signature error', error.errorCode || error.code || error.name || 'Unknown');
  } finally {
    event.completed();
  }
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
