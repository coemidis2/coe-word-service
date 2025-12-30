# COE MIDIS – Servicio Word (Flask)

Este servicio genera los Word (RP/RC) a partir de un JSON.

## Endpoints
- `GET /health`
- `POST /api/generar-word-rp`
- `POST /api/generar-word-rc`

## Seguridad opcional (X-API-KEY)
Si defines la variable de entorno `DOC_SERVICE_KEY`, el servicio exigirá el header:
- `X-API-KEY: <DOC_SERVICE_KEY>`

Si NO defines `DOC_SERVICE_KEY`, no valida nada (útil para pruebas).

## Deploy (Render/Railway)
- Build: `pip install -r requirements.txt`
- Start: `gunicorn app:app --bind 0.0.0.0:$PORT`

## Prueba rápida
```bash
curl -L -X POST "https://TU_DOMINIO/api/generar-word-rp" \
  -H "Content-Type: application/json" \
  -H "X-API-KEY: TU_KEY" \
  --data '{"codigo":"TEST-001","peligro":"Sismo","departamento":"Lima","distrito":"Miraflores"}' \
  --output test.docx
```
