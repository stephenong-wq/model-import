# Orion Tools

Two internal tools in one app:
- **Orion Import Builder** — parse Orion model exports, edit targets/bands on an interactive tree, export a filled import template, and reconcile an existing model export against a model-library target file.
- **Model Audit Tool** — compare a Master model export against a target file and flag overweight/underweight/missing holdings.

Switch between them with the nav bar at the top.

## Run locally

```bash
npm install
npm run dev
```

Opens at `http://localhost:5173`.

## Deploy — GitHub + Vercel

**1. Push this folder to a new GitHub repo:**

```bash
cd orion-tools
git init
git add .
git commit -m "Initial commit: Orion Import Builder + Model Audit Tool"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo-name>.git
git push -u origin main
```

(Create the empty repo on GitHub first — github.com → New repository — then use the URL it gives you.)

**2. Deploy on Vercel:**

- Go to [vercel.com/new](https://vercel.com/new) and sign in (GitHub login is easiest).
- Click **Import** next to the repo you just pushed.
- Vercel auto-detects Vite — the defaults (`npm run build`, output directory `dist`) are already correct. No config needed.
- Click **Deploy**.

That's it — you'll get a live URL, and every push to `main` auto-deploys from then on.

## Notes

- The Model Audit Tool's "remember my Master file" feature originally used Claude's artifact storage API. It's been swapped for a `localStorage`-backed equivalent (see `src/storageShim.js`) so it works the same way in a normal browser — saved data stays on whichever device/browser it was saved from, same as before.
- Both tools are lazy-loaded, so switching between them doesn't pull in both bundles upfront.
