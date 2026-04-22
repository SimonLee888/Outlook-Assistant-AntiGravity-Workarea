# Portfolio

Portfolio in Next.js, ready for Vercel.

## Local development

```bash
npm install
npm run dev
```

## Production build

```bash
npm run build
```

## Main files to edit

- Main homepage content: `components/portfolio-shell.tsx`
- Limited code access page: `components/code-lab.tsx`
- Shared content/data: `lib/portfolio-data.ts`
- Styles and animations: `components/portfolio-shell.module.css`
- Metadata: `app/layout.tsx`

## Vercel deployment

1. Push this folder to GitHub.
2. Import the repository in Vercel.
3. Framework preset: `Next.js`.
4. Build command: `npm run build`.
5. Output setting: automatic.

No environment variables are required.
