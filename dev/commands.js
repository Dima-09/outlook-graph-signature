function onNewMessageComposeHandler(event) {
  console.log('BDO: Event started');

  try {
    Office.context.mailbox.item.body.setSignatureAsync(
      'TEST',
      {
        coercionType: Office.CoercionType.Text,
      },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('BDO: Signature inserted');
        } else {
          console.error('BDO: Signature error', result.error);
        }

        event.completed();
      },
    );
  } catch (error) {
    console.error('BDO: Runtime error', error);
    event.completed();
  }
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
