/*
 * 在线传输下载 Service Worker。
 *
 * 拦截 __transfer-download/<key>/<filename> 请求，从同源 OPFS
 * （qtransfer/<key>）读出文件，以 attachment 形式响应。
 *
 * 为什么不用 blob URL：iOS Safari 对大体积 blob URL 的下载支持很差
 * （进度停滞、内存受限），而带 Content-Disposition 的真实 HTTP 响应
 * 才能触发 iOS 的「存储到文件」下载流程；File 主体由浏览器从磁盘
 * 流式读取，页面无需把整个文件再包一层 Blob。
 */

const DIR = 'qtransfer';
const MARK = '__transfer-download/';

self.addEventListener('install', () => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  if (event.request.method !== 'GET' || url.origin !== self.location.origin) return;
  const idx = url.pathname.indexOf(MARK);
  if (idx === -1) return;

  const rest = url.pathname.slice(idx + MARK.length);
  const slash = rest.indexOf('/');
  const key = slash >= 0 ? rest.slice(0, slash) : '';
  let name = '';
  try {
    name = slash >= 0 ? decodeURIComponent(rest.slice(slash + 1)) : '';
  } catch {
    name = '';
  }
  // key 由页面生成（session-counter），只允许安全字符
  if (!key || !/^[a-z0-9-]+$/i.test(key)) return;
  event.respondWith(serve(key, name));
});

async function serve(key, name) {
  try {
    const root = await navigator.storage.getDirectory();
    const dir = await root.getDirectoryHandle(DIR);
    const file = await (await dir.getFileHandle(key)).getFile();
    const safe = (name || 'file').replace(/[\\/:*?"<>|]/g, '_');
    return new Response(file, {
      headers: {
        'Content-Type': 'application/octet-stream',
        'Content-Length': String(file.size),
        'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(safe)}`,
        'Cache-Control': 'no-store',
      },
    });
  } catch {
    return new Response('file expired', { status: 404 });
  }
}
