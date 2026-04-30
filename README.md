# Matriz Comercial PX

Aplicacion web estatica para evaluar proyectos comerciales de Provexpress.

## Uso

Abre `index.html` en el navegador. La matriz calcula automaticamente:

- Costos directos del proyecto.
- Valor de venta antes de IVA.
- IVA y valor a facturar.
- Comision, impuestos, fletes, imprevistos y financiacion.
- Utilidad bruta y margenes Provexpress.

Los cambios se guardan en el navegador. Tambien puedes imprimir o exportar la tabla a un archivo `.xls`.

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

## Desarrollo local

Para correr la app con API local:

```powershell
npm run dev
```

Abre:

```text
http://localhost:3000
```

La primera consulta a Dataverse muestra un codigo de dispositivo en la terminal. Completa ese login con la cuenta corporativa y luego la app podra listar y crear matrices.
