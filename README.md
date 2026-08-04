# myrik

Web frontend for the order-report scripts.

## Local

```
python app.py            # http://127.0.0.1:8765
python app.py --selftest # checks
```

## Deploy (Vercel)

```
vercel        # preview
vercel --prod
```

`index.html` is the page, `api/index.py` wraps `app.py` as the function.
`vercel.json` declares that function explicitly: zero-config only picks up
`api/*.js`, it skips `api/*.py`, so the build fails without a `builds` entry.
Uploads run in-process and are thrown away after the response — nothing is
stored. Vercel caps request bodies at ~4.5 MB, so bigger exports need the
local server.
