# ARCHITECTURE.md — контекст и архитектура Outlook Graph Signature

## 1. Текущая ситуация

Репозиторий: https://github.com/Dima-09/outlook-graph-signature

Ветка: `main`. По итогам предыдущих обсуждений актуальная Graph-реализация находится в `dev/`, а в корне остались ресурсы первоначальной реализации. Действующий DEV опубликован по адресу `https://dima-09.github.io/outlook-graph-signature/dev/` и используется тестовыми учётными записями Microsoft 365. На момент подготовки задания пользователь сообщил об успешной централизованной установке и автоматической вставке подписи.

Ниже — карта по ранее изученной версии, а не свежий аудит рабочего дерева. Перед изменениями Codex обязан проверить фактические пути, импорты, манифест и зависимости. Если в репозитории уже есть изменения, учитывать их и не перезаписывать.

## 2. Карта существующих файлов

| Файл | Назначение / зависимость |
| --- | --- |
| `dev/manifest.xml` | XML-манифест Office Add-in: GUID, версия, разрешения, URL, события, desktop/mobile, команда Sign in. |
| `dev/graph-runtime.html` | HTML-среда событийного runtime, подключает Office.js с CDN Microsoft и `graph-runtime.js`. |
| `dev/graph-runtime-source.js` | Исходная точка входа автоматической подписи; импортирует `getUserProfile` и `buildSignature`. |
| `dev/graph-runtime.js` | Скомпилированный бандл событийной логики, используется Outlook. |
| `dev/graph-auth.js` | Создание NAA/MSAL-клиента, получение токена и профиля Graph. Импортирует `authconfig.js` и `@azure/msal-browser`. |
| `dev/authconfig.js` | Настройки Entra ID: Client ID, authority и любые прочие фактически присутствующие настройки. |
| `dev/signature-template.js` | Формирует HTML-подпись из объекта пользователя; не выполняет сетевые запросы. |
| `dev/signin.html` | Интерфейс панели Sign in, подключает Office.js и `signin.js`. |
| `dev/signin-source.js` | Обработчик кнопки Sign in, вызывает `getUserProfile(true)`, отображает данные или ошибку. |
| `dev/signin.js` | Готовый бандл панели Sign in. |
| `dev/commands.html`, `dev/commands.js` | Остаток промежуточной реализации. HTML может по-прежнему использоваться в SupportUrl и FormSettings. Проверить ссылки; старый статический обработчик не переносить в новую сборку без оснований. |
| `manifest.xml`, `commands.html`, `commands.js` в корне | Старая реализация. Не удалять без проверки старых установок и отдельного согласования. |
| `icon-16.png`, `icon-32.png`, `icon-80.png` в корне | Иконки, используемые в том числе актуальным DEV-манифестом. |
| `webpack.config.js` | Базовый конфиг, собирает `signin.js` из `signin-source.js`. |
| `webpack.graph.config.js` | Наследует базовый конфиг, меняет entry/output для `graph-runtime.js`. |
| `package.json`, `package-lock.json` | npm-зависимости, команды и зафиксированные версии. |
| `.gitignore`, `README.md` | Исключения Git и документация. |

Отдельный репозиторий `https://github.com/Dima-09/Dima-09.github.io` обслуживает корень текущего домена GitHub Pages. Там находится `.well-known/microsoft-officeaddins-allowed.json` с разрешённым URL событийного JavaScript. Этот репозиторий не входит в задание на редактирование.

## 3. Выполнение надстройки

```text
Установка XML в Microsoft 365 / Outlook
              │
      OnNewMessageCompose
              │
      ┌───────┴───────────┐
      │                   │
Classic Outlook       New / Web / Mobile
 JS runtime           HTML runtime
      │                   │
      │          graph-runtime.html + Office.js
      │                   │
      └───────┬───────────┘
              ▼
       graph-runtime.js
              │
    onNewMessageComposeHandler(event)
              │
       getUserProfile(false)
              │
      MSAL / NAA → Entra ID → token User.Read
              │
      Microsoft Graph /v1.0/me
              │
      buildSignature(user) → HTML
              │
      Office.js setSignatureAsync()
              │
      event.completed() в finally
```

В Classic Outlook указанный в манифесте JavaScript override позволяет использовать JS runtime; HTML-страница служит средой выполнения для остальных поддерживаемых клиентов. Конкретный механизм зависит от клиента и его поддержки Office API. Сохранить существующие Runtime, LaunchEvent, Office.actions.associate, разрешения и требования манифеста.

Панель Sign in — отдельная точка входа: `manifest.xml` → `signin.html` → `signin.js` → `getUserProfile(true)`. Она предназначена для диагностики подключения и интерактивной авторизации при необходимости, не вставляет подпись сама. Автоматический обработчик сначала пытается получить токен бесшумно и не должен открывать popup.

Graph использует делегированное `User.Read`, endpoint `/v1.0/me` и текущий `$select`: `displayName,jobTitle,department,companyName,mail,userPrincipalName,businessPhones,mobilePhone`. Шаблон использует персональные данные, телефоны компаний, правила внутреннего номера и юридический текст. Сохранить всё поведение, включая fallback, нормализацию, экранирование HTML и форматирование. Даже неиспользуемые поля запроса не оптимизировать в рамках структурного рефакторинга.

## 4. Сборка и зависимости

```text
                         Webpack
                            │
       graph-runtime-source.js     signin-source.js
              │    │                       │
              │    └──── graph-auth.js ────┘
              │                 ├── authconfig.js
              │                 └── @azure/msal-browser
              └── signature-template.js
                     │                      │
             graph-runtime.js            signin.js
```

Это зависимости исходников при сборке. Outlook не скачивает эти исходные модули по отдельности: нужный код уже включён в бандлы. Office.js остаётся внешним API хоста/CDN. Существующие параметры single-chunk и отсутствие разделяемого runtime важны для событийного JS runtime. В ходе рефакторинга сохранить эквивалентные настройки.

## 5. Целевая архитектура

```text
outlook-graph-signature/
├── AGENTS.md
├── ARCHITECTURE.md
├── REFACTORING.md
├── README.md
├── src/
│   ├── authconfig.js
│   ├── graph-auth.js
│   ├── graph-runtime-source.js
│   ├── signature-template.js
│   └── signin-source.js
├── config/
│   ├── dev.json
│   └── prod.json
├── public/
│   ├── graph-runtime.html
│   ├── signin.html
│   ├── commands.html
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
├── manifests/
│   └── manifest.template.xml
├── scripts/
│   └── build.cjs
├── webpack.config.js
├── package.json
├── package-lock.json
├── .gitignore
└── dist/                         # генерируется, не в Git
    ├── dev/
    └── prod/
```

Схема отражает целевые роли файлов. Названия исходников можно сохранить для минимального diff. Точная реализация скрипта, тестов и шаблонных переменных — на усмотрение Codex после согласования плана. На переходном этапе действующие опубликованные `dev/` и корневые ресурсы остаются дополнительно к этой структуре.

Каждая папка сборки содержит самостоятельный комплект:

```text
dist/<env>/
├── .well-known/
│   └── microsoft-officeaddins-allowed.json
├── manifest.xml
├── graph-runtime.html
├── graph-runtime.js
├── signin.html
├── signin.js
├── commands.html
├── icon-16.png
├── icon-32.png
└── icon-80.png
```

Сохранять необходимые дополнительные ресурсы и лицензионные уведомления Webpack, если они создаются. XML нужен для установки, но не обязательно должен находиться на сайте после установки. Общие ресурсы копируются при сборке; генерируемые параметры подставляются автоматически.

## 6. Разделение сред и хостинга

| Параметр | DEV | PROD |
| --- | --- | --- |
| Исходники | Общие | Общие |
| Ветка сейчас | main | main |
| Локальный вывод | `dist/dev/` | `dist/prod/` |
| Команда | `npm run build:dev` | `npm run build:prod` |
| Надстройка Outlook | Отдельный GUID | Другой GUID |
| Entra ID | DEV App registration | Отдельная PROD App registration (ещё не настроена) |
| Хостинг (план) | Отдельный Cloudflare Pages | Отдельный Cloudflare Pages |
| Публикация | Вручную | Вручную после проверки DEV |

На двух независимых сайтах содержимое `dist/dev/` и `dist/prod/` публикуется в корни соответствующих сайтов. Локальные `dist`, `dev` и `prod` не обязаны входить в публичные URL. Будущий URL может быть `https://<dev-host>/graph-runtime.js` или `https://<prod-host>/graph-runtime.js` (условные примеры, не реальные адреса).

У каждого сайта `.well-known/microsoft-officeaddins-allowed.json` должен быть в корне origin. Его содержимое включает точный публичный URL событийного JS из манифеста. При одном общем origin потребуется единый корневой JSON с обоими разрешёнными URL, а не два файла в подпапках. Сейчас целевой сценарий — два отдельных origin.

Переезд на новый origin требует согласованных обновлений манифеста, Entra SPA redirect URI для NAA (`brk-multihub://<origin-host>` с фактическими требованиями Microsoft), `.well-known` и последующей проверки в клиентах Outlook. Это отдельный этап, не задача текущего рефакторинга.
