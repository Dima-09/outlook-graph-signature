import { getUserProfile } from './graph-auth.js';

function setSignature(text) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSignatureAsync(
      text,
      {
        coercionType: Office.CoercionType.Text,
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
    const user = await getUserProfile(false);

    const signature = ['GRAPH OK', user.displayName, user.jobTitle, user.department, user.mail, ...(user.businessPhones || []), user.mobilePhone].filter(Boolean).join('\n');

    await setSignature(signature);
  } catch (error) {
    console.log('Graph error:', error);

    try {
      await setSignature('GRAPH ERROR: ' + (error.errorCode || error.code || error.name || 'Unknown'));
    } catch (signatureError) {
      console.log(signatureError);
    }
  } finally {
    event.completed();
  }
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
