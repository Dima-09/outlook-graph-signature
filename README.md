# Outlook Graph Signature

Корпоративная надстройка BDO вставляет персональную HTML-подпись при создании письма в Outlook. Данные сотрудника поступают из Microsoft Graph. Панель **Sign in** в настольном интерфейсе позволяет выполнить интерактивный вход и проверить полученный профиль.

Исходники и история разработки находятся в [GitHub](https://github.com/Dima-09/outlook-graph-signature). DEV и PROD собираются из общей ветки `main` и размещаются вручную в двух отдельных Cloudflare Workers Static Assets.

## Быстрый старт

Для локальной работы нужны Git, Node.js 22.12 или новее и npm. Выполните из каталога разработки:

```sh
git clone https://github.com/Dima-09/outlook-graph-signature.git
cd outlook-graph-signature
npm ci
npm test
npm run build:all
```

В PowerShell при блокировке `npm.ps1` используйте `npm.cmd`. Для сборки одной среды есть `npm run build:dev` и `npm run build:prod`. В `dist/dev/` и `dist/prod/` появятся независимые комплекты с собственными `manifest.xml`. Для локальных тестов и компиляции достаточно установленных зависимостей.

**Перед публикацией проверьте конфигурацию.** В этой копии обе среды помечены `configured: false`, а PROD содержит тестовые значения. Предупреждение `NOT CONFIGURED — DO NOT DEPLOY` означает, что комплект предназначен для проверки сборки.

## Документация

- [DEPLOYMENT.md](DEPLOYMENT.md) — подготовка инфраструктуры, выпуск, испытания в Outlook и откат.
- [ARCHITECTURE.md](ARCHITECTURE.md) — исходники, путь данных, контракт сборки и манифест.
- [AGENTS.md](AGENTS.md) — правила изменений и приёмки.
