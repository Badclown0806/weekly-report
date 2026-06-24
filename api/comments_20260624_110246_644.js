// Vercel Serverless Function — 评论 CRUD via GitHub Contents API
// 部署后路径: /api/comments

const GITHUB_TOKEN = process.env.GITHUB_PAT || '';
if (!GITHUB_TOKEN) {
  console.error('Missing GITHUB_PAT environment variable');
}
const REPO_OWNER = 'Badclown0806';
const REPO_NAME = 'weekly-report';
const FILE_PATH = 'comments.json';
const API_BASE = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;

const HEADERS = {
  'Authorization': `Bearer ${GITHUB_TOKEN}`,
  'Accept': 'application/vnd.github.v3+json',
  'User-Agent': 'vercel-serverless-comments'
};

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PATCH, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    }
  });
}

// ── 编解码：兼容客户端 btoa(unescape(encodeURIComponent(str))) ──
// 客户端使用 btoa(unescape(encodeURIComponent(str))) 将 UTF-8 中文转为 base64，
// 此编码方案等价于：UTF-8 bytes → 每个 byte 作为 Latin-1 字符 → base64。
// Node.js 端解码：base64 → bytes → 构造 %xx 格式 → decodeURIComponent。

function decodeClientBase64(b64) {
  const bytes = Buffer.from(b64, 'base64');
  const percentEncoded = Array.from(bytes)
    .map(b => '%' + b.toString(16).padStart(2, '0').toUpperCase())
    .join('');
  return decodeURIComponent(percentEncoded);
}

function encodeClientBase64(str) {
  const encoded = encodeURIComponent(str);
  let result = '';
  for (let i = 0; i < encoded.length; i++) {
    const c = encoded.charAt(i);
    if (c === '%') {
      const hex = encoded.substring(i + 1, i + 3);
      result += String.fromCharCode(parseInt(hex, 16));
      i += 2;
    } else {
      result += c;
    }
  }
  return Buffer.from(result, 'binary').toString('base64');
}

async function getComments() {
  const res = await fetch(API_BASE, { headers: HEADERS });
  if (!res.ok) {
    if (res.status === 404) return [];
    throw new Error(`GitHub GET failed: ${res.status} ${await res.text()}`);
  }
  const data = await res.json();
  const content = decodeClientBase64(data.content);
  return { comments: JSON.parse(content), sha: data.sha };
}

async function saveComments(comments, sha, message) {
  const content = encodeClientBase64(JSON.stringify(comments, null, 2));
  const body = { message, content, sha };
  const res = await fetch(API_BASE, {
    method: 'PUT',
    headers: { ...HEADERS, 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  if (!res.ok) {
    throw new Error(`GitHub PUT failed: ${res.status} ${await res.text()}`);
  }
  return res.json();
}

async function handler(req) {
  try {
    // CORS preflight
    if (req.method === 'OPTIONS') {
      return new Response(null, {
        status: 204,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, POST, PATCH, DELETE, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type'
        }
      });
    }

    // GET — 返回所有评论
    if (req.method === 'GET') {
      const { comments } = await getComments();
      return jsonResponse(comments);
    }

    // POST — 新增评论
    if (req.method === 'POST') {
      const body = await req.json();
      const { comments, sha } = await getComments();
      const newComment = {
        id: Date.now().toString(36) + Math.random().toString(36).slice(2, 8),
        date: body.date || new Date().toISOString().slice(0, 10),
        week: body.week || '',
        sku: body.sku || '',
        shop: body.shop || '',
        text: body.text || '',
        author: body.author || '匿名',
        created_at: new Date().toISOString()
      };
      comments.push(newComment);
      await saveComments(comments, sha, `Add comment ${newComment.id}`);
      return jsonResponse(newComment, 201);
    }

    // PATCH — 编辑评论
    if (req.method === 'PATCH') {
      const body = await req.json();
      const { comments, sha } = await getComments();
      const idx = comments.findIndex(c => c.id === body.id);
      if (idx === -1) {
        return jsonResponse({ error: 'Comment not found' }, 404);
      }
      if (body.text !== undefined) comments[idx].text = body.text;
      comments[idx].edited_at = new Date().toISOString();
      await saveComments(comments, sha, `Edit comment ${body.id}`);
      return jsonResponse(comments[idx]);
    }

    // DELETE — 删除评论
    if (req.method === 'DELETE') {
      const body = await req.json();
      const { comments, sha } = await getComments();
      const idx = comments.findIndex(c => c.id === body.id);
      if (idx === -1) {
        return jsonResponse({ error: 'Comment not found' }, 404);
      }
      const removed = comments.splice(idx, 1)[0];
      await saveComments(comments, sha, `Delete comment ${body.id}`);
      return jsonResponse(removed);
    }

    return jsonResponse({ error: 'Method not allowed' }, 405);
  } catch (err) {
    console.error('[comments API]', err);
    return jsonResponse({ error: err.message }, 500);
  }
}

// Vercel Node.js runtime exports
module.exports = handler;
