# Matriz Comercial PX

Aplicacion web estatica para evaluar proyectos comerciales de Provexpress.

## Uso

Publica o abre la app desde un origen web permitido. La matriz calcula automaticamente:

- Costos directos del proyecto.
- Valor de venta antes de IVA.
- IVA y valor a facturar.
- Comision, impuestos, fletes, imprevistos y financiacion.
- Utilidad bruta y margenes Provexpress.
- Seguimiento de fase del producto: Solicitud de credito, Oferta y Cierre.

Los cambios se guardan localmente mientras se edita. Con Microsoft 365 conectado tambien puede guardar y listar matrices en Dataverse. Tambien puedes imprimir o exportar la tabla a un archivo `.xls`.

## Dataverse

Entorno de desarrollo:

```text
https://db-px.crm2.dynamics.com/
```

La definicion inicial de tablas y columnas esta en `dataverse/schema.json`.

App registration de desarrollo:

```text
ffcf61c2-18a2-4be4-a753-c81934026e4d
```

Para autenticar la CLI de Power Platform contra el entorno:

```powershell
pac auth create --environment https://db-px.crm2.dynamics.com/
```

Para revisar el esquema que se va a crear sin tocar Dataverse:

```powershell
npm run dataverse:plan
```

Para aplicar el esquema:

```powershell
npm run dataverse:setup
```

El script abre un inicio de sesion por codigo de dispositivo. La App Registration debe tener permiso delegado de Dataverse `user_impersonation`.

## Publicacion

La app esta pensada para GitHub Pages: no requiere backend Node para operar. El navegador usa MSAL y llama directamente la Dataverse Web API con la sesion del usuario.

La App Registration debe tener estos redirect URI como SPA, segun donde se publique:

```text
https://provexpress.github.io/Matriz-Comercial-PX/
http://localhost:5500/
```

Tambien debe tener permiso delegado:

```text
Dynamics CRM / Dataverse: user_impersonation
```

Para probar localmente, usa un servidor estatico, no abras el archivo con `file://`. Por ejemplo:

```powershell
npx serve . -l 5500
```

El consecutivo se calcula consultando el ultimo registro `MX-AAAA-0000`. Para uso intensivo conviene cambiar `Consecutivo` a autonumeracion de Dataverse y evitar duplicados si dos usuarios guardan al mismo tiempo.
