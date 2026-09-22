# Meesho Label Sorter

A browser-only tool that:

- reads an original Meesho shipping-label PDF;
- sorts label pages in the exact size order `XXS, XS, S, M, L, XL, XXL, 3XL, 4XL, 5XL, 6XL, 7XL, FREE SIZE` and then by SKU;
- preserves the original label content and adds the selected date in the blank area;
- downloads the sorted PDF and a size/SKU-wise Excel picklist with order counts.

PDF processing happens locally in the browser. The uploaded file is not sent to a server.

## Local development

```bash
npm install
npm run dev
```

## GitHub Pages

The included GitHub Actions workflow builds and publishes the site on every push to `main`.

In the repository, open **Settings → Pages** and select **GitHub Actions** as the source. The Vite base path is relative, so the same build works for both project Pages and a custom domain.
