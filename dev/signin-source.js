import { getUserProfile } from './graph-auth.js';

Office.onReady(function () {
  const button = document.getElementById('signin');
  const result = document.getElementById('result');

  button.disabled = false;

  button.onclick = async function () {
    button.disabled = true;
    result.textContent = 'Connecting to Microsoft Graph...';

    try {
      const user = await getUserProfile(true);

      result.textContent = JSON.stringify(user, null, 2);
    } catch (error) {
      result.textContent = 'ERROR:\n' + (error.errorCode || error.code || '') + '\n' + error.message;
    } finally {
      button.disabled = false;
    }
  };
});
