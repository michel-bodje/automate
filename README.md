# Automate for Allen Madelin

An Office Add-in for Allen Madelin law firm.

## Setup locally (development)
1. Clone this repo.
2. Run `npm install`.
3. Run `npx office-addin-dev-settings m365-account login` and select your Office work or school account.
3. Run `npm start` to test on your desktop.

## Setup company-wide (production)
This assumes that all targeted computers have access to the same Office account.

1. Fork this repo, and enable GitHub Pages (or your preferred hosting method).
2. Verify `manifests/` files contain a link to the [GitHub Pages](https://michel-bodje.github.io/automate/).
2. Run `npm run build`. This will generate the build files in `docs/`.
3. Commit the build files to your fork.
4. Wait for deployment.
5. Enjoy!

Read the [user manual](user-manual.md).
