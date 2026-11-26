function normalize(value) {
  return typeof value === 'string' ? value.trim() : '';
}

function escapeHtml(value) {
  return normalize(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

export function buildSignature(user) {
  const name = escapeHtml(normalize(user.displayName).toUpperCase());

  const jobTitle = escapeHtml(user.jobTitle);

  const companyName = normalize(user.department) || 'BDO Qazaqstan';
  const company = escapeHtml(companyName);

  const companyPhones = {
    'bdo qazaqstan': '+7 727 331 31 34',
    'bdo outsourcing qazaqstan': '+7 727 331 31 37',
  };

  const companyKey = companyName.toLowerCase().replace(/\s+/g, ' ');

  const officePhone = companyPhones[companyKey] || companyPhones['bdo qazaqstan'];

  const extension = escapeHtml(user.businessPhones?.[0]);

  const directPhone = extension ? `Direct: ${officePhone} ext. ${extension}` : '';

  const mobilePhone = escapeHtml(user.mobilePhone);
  const email = normalize(user.mail);

  const emailHtml = email
    ? /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(email)
      ? `<a href="mailto:${encodeURIComponent(email).replace(/%40/g, '@')}"
                 style="color: #e81a3b; text-decoration: underline;">
                 ${escapeHtml(email)}
               </a>`
      : escapeHtml(email)
    : '';

  const personalLines = [name ? `<strong>${name}</strong>` : '', jobTitle, company, directPhone, mobilePhone ? `Mobile: ${mobilePhone}` : '', emailHtml].filter(Boolean).join('<br>');

  return `
        <div style="
            font-family: 'Trebuchet MS', Arial, sans-serif;
            font-size: 10pt;
            color: #333333;
            line-height: 13.5pt;
            text-align: left;
        ">
            ${personalLines}

            <br><br>

            6A Gabdullin Street<br>
            A15H4E3, Almaty<br>
            KAZAKHSTAN<br>
            Tel: ${officePhone}<br>

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
                ${company}, a limited liability partnership,
                registered under the laws of the Republic of Kazakhstan,
                is a member of BDO International Limited, a UK company
                limited by guarantee, and forms part of the international
                BDO network of independent member firms. BDO is the brand
                name for the BDO network and for each of the BDO Member Firms.
            </div>
        </div>
    `;
}
