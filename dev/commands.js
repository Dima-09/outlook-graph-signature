function buildSignature() {
  return `
        <div style="
            font-family: 'Trebuchet MS', Arial, sans-serif;
            font-size: 10pt;
            color: #333333;
            line-height: 13.5pt;
            text-align: left;
        ">
            <strong>DMITRY SMIRNOV</strong><br>
            IT<br>
            BDO Qazaqstan<br>
            Direct: +7 727 331 31 34 ext. 999<br>
            Mobile: +71234567788<br>

            <a href="mailto:d.smirnov@bdoqz.com"
              style="color: #e81a3b; text-decoration: underline;">
                d.smirnov@bdoqz.com
            </a>

            <br><br>

            6A Gabdullin Street<br>
            A15H4E3, Almaty<br>
            KAZAKHSTAN<br>
            Tel: +7 727 331 31 34<br>

            <a href="https://www.bdoqz.com"
              style="color: #e81a3b; text-decoration: underline;">
                www.bdoqz.com
            </a>

            <br><br>

            <div style="
                font-family: 'Trebuchet MS', Arial, sans-serif;
                font-size: 8pt;
                color: #333333;
                line-height: 10.5pt;
                max-width: 420px;
            ">
                BDO Qazaqstan, a limited liability partnership,
                registered under the laws of the Republic of Kazakhstan,
                is a member of BDO International Limited, a UK company
                limited by guarantee, and forms part of the international
                BDO network of independent member firms. BDO is the brand
                name for the BDO network and for each of the BDO Member Firms.
            </div>
        </div>
    `;
}

function onNewMessageComposeHandler(event) {
  console.log('BDO: Event started');

  try {
    var signature = buildSignature();

    Office.context.mailbox.item.body.setSignatureAsync(
      signature,
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
