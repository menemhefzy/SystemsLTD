# AGENTS.md

## Cursor Cloud specific instructions

This is a **static content repository** (no build tools, package managers, tests, or linters). It contains marketing/website content for Mastplast — an "About Us" page in Markdown (`mastplast_about_us.md`) and as an Odoo 18 HTML snippet (`mastplast_about_us_odoo.html`).

### Serving content locally

To preview the HTML content, use Python's built-in HTTP server:

```
python3 -m http.server 8080 --directory /workspace
```

Then open `http://localhost:8080/mastplast_about_us_odoo.html` in a browser. The HTML file uses Bootstrap CSS classes but does not include Bootstrap via CDN — it is designed to be embedded in an Odoo 18 website builder which provides Bootstrap.

### Key notes

- There are **no dependencies to install**, no `package.json`, `requirements.txt`, or any other manifest.
- There are **no automated tests or linters** configured.
- The HTML file (`mastplast_about_us_odoo.html`) is an Odoo 18 snippet — it relies on Bootstrap classes being provided by the Odoo framework.
- The update script is a no-op (`echo "No dependencies to install"`) since there is nothing to install.
