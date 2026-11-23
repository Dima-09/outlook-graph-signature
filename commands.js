function onNewMessageComposeHandler(event) {
  console.log('BDO: Compose event started');

  try {
    Office.context.mailbox.item.body.setSignatureAsync(
      '<div><strong>TEST SIGNATURE</strong></div>',
      {
        coercionType: Office.CoercionType.Html,
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

// Если в установленном XML осталось второе событие,
// оно не должно повторно вставлять подпись.
Office.actions.associate('onMessageRecipientsChangedHandler', function (event) {
  event.completed();
});
