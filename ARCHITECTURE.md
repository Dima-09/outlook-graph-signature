# Архитектура Outlook Graph Signature

Outlook загружает клиентское приложение по HTTPS. Office.js предоставляет контекст письма, MSAL Browser получает токен через Nested App Authentication (NAA), Microsoft Graph возвращает профиль. На Cloudflare работают два независимых Workers со статическими ресурсами. GitHub хранит исходники.

## Структура репозитория

| Путь | Назначение |
| --- | --- |
| `src/graph-runtime-source.js` | Регистрация обработчика Office и вставка подписи |
| `src/signin-source.js` | Панель интерактивного входа |
| `src/graph-auth.js`, `src/authconfig.js` | Получение токена, запрос профиля и параметры MSAL |
| `src/signature-template.js` | Экранирование полей, оформление и текст подписи |
| `config/dev.json`, `config/prod.json` | Публичные параметры сред |
| `public/` | HTML событийного runtime, панели и страницы commands, иконки 16/32/80 |
| `manifests/manifest.template.xml` | Общий XML-шаблон Outlook |
| `scripts/build.cjs`, `webpack.config.js` | Валидация, две компиляции Webpack и упаковка результата |
| `tests/build.test.cjs`, `tests/signature.test.cjs` | Проверки сборщика и подписи на фиктивных профилях |
| `package.json`, `package-lock.json`, `.gitignore` | Команды, дерево зависимостей и исключения Git |
| `dist/dev/`, `dist/prod/` | Генерируемые комплекты |

## От создания письма до вставки подписи

1. Outlook активирует `OnNewMessageCompose`. Манифест связывает его с `onNewMessageComposeHandler`, зарегистрированным через `Office.actions.associate`.
2. `getUserProfile(false)` запрашивает токен с делегированным `User.Read`. MSAL создаёт nestable public client и вызывает `acquireTokenSilent`. В панели Sign in используется `getUserProfile(true)`, где после `InteractionRequiredAuthError` допускается `acquireTokenPopup`.
3. Клиент отправляет Bearer-токен в `GET https://graph.microsoft.com/v1.0/me`. Параметр `$select` перечисляет `displayName,jobTitle,department,companyName,mail,userPrincipalName,businessPhones,mobilePhone`. Панель показывает ответ или ошибку. Событийный обработчик требует непустой `displayName`.
4. Шаблон экранирует пользовательские строки, берёт компанию из `department` с резервным значением BDO Qazaqstan, добавочный номер из первого `businessPhones`, адрес почты из `mail`. Обработчик вставляет HTML через `setSignatureAsync`. При ошибке он записывает её код или имя и завершает событие вызовом `event.completed()` в `finally`.

Манифест запрашивает у Outlook `ReadWriteItem`. Базовый Mailbox — 1.1, в VersionOverrides — 1.5. Для Classic Outlook предусмотрен JavaScript override. Браузерный runtime загружает HTML и Office.js. Событие объявлено в desktop и mobile, кнопка Sign in — только в DesktopFormFactor. Страница `commands.html` используется в SupportUrl и FormSettings. Требования событий, NAA и версий клиентов проверяются при приёмке по [DEPLOYMENT.md](DEPLOYMENT.md).

## Контракт сборки

Сборщик проверяет типы, обязательные поля, UUID, HTTPS URL и версию манифеста. При `configured: true` он отклоняет домены `.invalid` и известные тестовые UUID. При `false` он разрешает сборку с предупреждением. Список полей и порядок их заполнения приведены в [DEPLOYMENT.md](DEPLOYMENT.md).

Webpack в режиме production подставляет Client ID и authority через DefinePlugin и последовательно создаёт два самодостаточных файла — `graph-runtime.js` и `signin.js`. Сборщик копирует публичные ресурсы, сохраняет созданные Webpack лицензии, подставляет экранированные значения в XML и записывает URL событийного JS в `.well-known/microsoft-officeaddins-allowed.json`. Он проверяет состав комплекта и ссылки до замены выходного каталога. Публичные адреса вычисляются от `baseUrl`.

Каждая среда имеет отдельную блокировку и временный каталог. При замене результата сборщик временно сохраняет предыдущий каталог и пытается восстановить его при ошибке. Ссылки и junction в управляемых каталогах отклоняются. При аварийной остановке проверьте отсутствие активного сборщика и содержимое `dist/.backup-<target>-*` перед удалением оставшейся `.lock-<target>`. Команда `build:all` выполняет DEV и PROD последовательно. При ошибке PROD уже собранный DEV сохраняется.
