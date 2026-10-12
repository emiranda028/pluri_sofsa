# Sitio web LTELC · ltelc.tech

Sitio estático de una sola página (`index.html`), sin build. Usa Three.js (escena 3D de la portada), GSAP + ScrollTrigger (animaciones de scroll) y Lenis (scroll suave), cargados desde CDN. Si alguno no carga, la página sigue funcionando sin animaciones.

Capturas de sistemas: reemplazar las maquetas de "Sistema de uso interno" agregando la imagen en `img/` y el campo `img` del proyecto en la lista `PROJECTS` del script.

| Archivo | Para qué |
|---|---|
| `index.html` | El sitio completo (estilos y script incluidos) |
| `favicon.svg` | Ícono de la pestaña |
| `img/` | Fotos (Unsplash, uso comercial libre) y capturas de proyectos públicos |
| `CNAME` | Dominio para GitHub Pages (`ltelc.tech`) |
| `robots.txt`, `sitemap.xml` | Indexación en buscadores |

Datos de contacto: se editan en el bloque `CONTACT` del script al final de `index.html` y en la sección `#contacto`.

## Publicar en ltelc.tech

Hoy el dominio apunta a los nameservers de *parking* de Hostinger (`lunar/solar.dns-parking.com`), así que muestra la página de estacionamiento. Hay dos caminos:

### A) Hosting de Hostinger
1. Contratar un plan de Web Hosting y asociarle `ltelc.tech` (Hostinger cambia los DNS solo).
2. hPanel → Sitios web → Administrar → **Administrador de archivos** → carpeta `public_html`.
3. Subir `index.html`, `favicon.svg`, `robots.txt` y `sitemap.xml` (el `CNAME` no hace falta).
4. Activar SSL gratis en hPanel → Seguridad → SSL.

### B) GitHub Pages (gratis)
1. Crear un repositorio público (por ejemplo `ltelc-web`) con el contenido de esta carpeta en la raíz.
2. Settings → Pages → Deploy from branch → `main` / root. Custom domain: `ltelc.tech`, marcar *Enforce HTTPS*.
3. En Hostinger → Dominios → ltelc.tech → **DNS / Nameservers** → registros DNS:
   - Borrar los registros `A` y `CNAME` de parking existentes para `@` y `www`.
   - `A` `@` → `185.199.108.153`, `185.199.109.153`, `185.199.110.153`, `185.199.111.153`
   - `CNAME` `www` → `<usuario-github>.github.io`
4. Esperar la propagación (de minutos a unas horas).

## Contacto
El sitio usa `agencialtelc@gmail.com` y el WhatsApp +54 9 11 5706-2947 (bloque `CONTACT` en `index.html`).
