# myrik

Web frontend for the order-report scripts.

## Local

```
pip install -r requirements.txt
python app.py             # http://127.0.0.1:8765
python app.py --selftest  # checks
```

## Vercel

```
vercel deploy
```

`api/index.py` is the entrypoint; `vercel.json` routes every path to it and
bundles the report scripts alongside. Uploads are capped at 4.5 MB (Vercel
request-body limit) and each run must finish inside `maxDuration` (60s).
