# GhostLayer

GhostLayer is a zero-backend browser experiment that reveals recoverable state hiding inside ordinary PowerPoint files.

Drop a `.pptx` and GhostLayer locally inspects the OOXML package for:

- cropped images whose original embedded pixels remain recoverable;
- slides marked hidden from slideshow mode;
- presenter notes that travel with the file;
- shapes positioned fully outside the visible slide canvas;
- document author and last-editor metadata.

The built-in synthetic board-deck specimen is generated in the browser and passed through the same ZIP/PPTX scanner as uploaded files. Uploaded presentations never leave the browser.

## Run

Serve the repository as static files and open `index.html`. No API key, account, backend, database, or build step is required.

## Verify

```bash
npm test
```

The core parser is deterministic. It does not infer hidden intent; it compares visible presentation state with structurally present content inside the PPTX package.
