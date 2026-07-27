# Polaria UI Runner (tercero, fuera del repo)

Este paquete corre **por UI del front** y no modifica tu código de `polaria-wms-web`.

## Qué hace

- Usa usuario configurador para crear:
  - empresa
  - cuenta
  - bodega
  - usuarios por rol
- Entra con cada administrador de cuenta y:
  - crea proveedor, cliente, comprador, camión
  - importa el Excel de su simulación
- Ejecuta ciclo A y ciclo B de:
  - Andino (Chrome)
  - Mar Azul (Edge)
  - Aves (Opera)
- Corre todo en paralelo y muestra monitoreo en consola.

## Requisitos

1. Front corriendo en `http://localhost:3001`
2. Back corriendo (lo levantas tú)
3. Windows con rutas:
   - Opera: `C:\Program Files\Opera GX\opera.exe`
   - Chrome: `C:\Program Files\Google\Chrome\Application\chrome.exe`
   - Edge: `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`
4. Excel en Descargas:
   - `Andino(1).xlsx`
   - `Mar_Azul(1).xlsx`
   - `Aves_Dorada(1).xlsx`

## Instalación

```bash
cd polaria-ui-runner
npm install
```

## Ejecución

```bash
node run.mjs --downloads "C:\Users\TU_USUARIO\Downloads" --front-url "http://localhost:3001" --headed
```

## Nota importante

El runner está basado en selectores UI y puede requerir ajuste fino si cambian textos/botones del front.
# simualciones
