// Cloudflare Pages Function: /api/bible
// 갓피아(godpia.com) 성경 본문 스크래핑 프록시
// 참고: https://github.com/Seongmin3927/biblereference_godsitepia
//
// 브라우저에서 godpia.com 으로 직접 요청하면 CORS 정책에 막히므로,
// 이 서버리스 함수가 대신 요청하여 한 장(chapter) 전체 구절을 JSON 으로 돌려줍니다.
//
// 요청: GET /api/bible?vol=act&chap=1&ver=gae
// 응답: { success: true, data: { vol, chapter, version, verses: [{ verse, text }] } }

// 갓피아 책 코드 → 장 수 (유효성 검증용)
const BOOK_CHAPTERS = {
  gen: 50, exo: 40, lev: 27, num: 36, deu: 34, jos: 24, jdg: 21, rut: 4,
  '1sa': 31, '2sa': 24, '1ki': 22, '2ki': 25, '1ch': 29, '2ch': 36,
  ezr: 10, neh: 13, est: 10, job: 42, psa: 150, pro: 31, ecc: 12, sng: 8,
  isa: 66, jer: 52, lam: 5, ezk: 48, dan: 12, hos: 14, jol: 3, amo: 9,
  oba: 1, jon: 4, mic: 7, nam: 3, hab: 3, zep: 3, hag: 2, zec: 14, mal: 4,
  mat: 28, mrk: 16, luk: 24, jhn: 21, act: 28, rom: 16, '1co': 16, '2co': 13,
  gal: 6, eph: 6, php: 4, col: 4, '1th': 5, '2th': 3, '1ti': 6, '2ti': 4,
  tit: 3, phm: 1, heb: 13, jas: 5, '1pe': 5, '2pe': 3, '1jn': 5, '2jn': 1,
  '3jn': 1, jud: 1, rev: 22
};

const BIBLE_VERSIONS = {
  gae: '개역개정4판',
  han: '개역한글',
  saenew: '새번역',
  easy: '쉬운성경',
  hyun: '현대인의성경',
  niv: 'NIV'
};

const JSON_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Content-Type': 'application/json; charset=utf-8'
};

function json(body, status = 200, extraHeaders = {}) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...JSON_HEADERS, ...extraHeaders }
  });
}

function decodeEntities(str) {
  return str
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n, 10)))
    .replace(/&amp;/g, '&');
}

// 갓피아 HTML 에서 <li dataSec="N"> ... </li> 단위로 절 번호와 본문 추출
export function parseGodpiaHtml(html) {
  const liRegex = /<li\b[^>]*>([\s\S]*?)<\/li>/gi;
  const verses = [];
  let match;

  while ((match = liRegex.exec(html)) !== null) {
    const fullLi = match[0];
    const liInner = match[1];

    let verseNum = null;
    const dataSecMatch = fullLi.match(/dataSec=["'](\d+)["']/i);
    if (dataSecMatch) {
      verseNum = parseInt(dataSecMatch[1], 10);
    } else {
      const noMatch = liInner.match(/<span[^>]*class=["'][^"']*bible-read-no[^"']*["'][^>]*>\s*(\d+)\s*<\/span>/i);
      if (noMatch) verseNum = parseInt(noMatch[1], 10);
    }
    if (verseNum === null) continue;

    // 절 번호 span 제거 → 나머지 태그 제거 → 엔티티/공백 정리
    let text = liInner.replace(/<span[^>]*class=["'][^"']*bible-read-no[^"']*["'][^>]*>[\s\S]*?<\/span>/gi, '');
    text = decodeEntities(text.replace(/<[^>]+>/g, '')).replace(/\s+/g, ' ').trim();

    if (text) verses.push({ verse: verseNum, text });
  }

  return verses;
}

export async function onRequest(context) {
  const { request } = context;

  if (request.method === 'OPTIONS') {
    return new Response(null, { status: 204, headers: JSON_HEADERS });
  }

  try {
    const url = new URL(request.url);
    const vol = (url.searchParams.get('vol') || '').toLowerCase();
    const chap = parseInt(url.searchParams.get('chap') || '', 10);
    const ver = BIBLE_VERSIONS[url.searchParams.get('ver')] ? url.searchParams.get('ver') : 'gae';

    if (!BOOK_CHAPTERS[vol]) {
      return json({ success: false, error: `성경 코드(${vol || '없음'})가 올바르지 않습니다.` }, 400);
    }
    if (isNaN(chap) || chap < 1 || chap > BOOK_CHAPTERS[vol]) {
      return json({ success: false, error: `장 번호가 올바르지 않습니다. (1~${BOOK_CHAPTERS[vol]}장)` }, 400);
    }

    // Cloudflare 엣지 캐시: 같은 장을 반복 요청할 때 갓피아 서버 부하를 줄입니다.
    const cache = typeof caches !== 'undefined' ? caches.default : null;
    const cacheKey = new Request(`https://godpia-cache.local/${ver}/${vol}/${chap}`);
    if (cache) {
      const cached = await cache.match(cacheKey);
      if (cached) return cached;
    }

    const postBody = new URLSearchParams({
      ver,
      ver2: '',
      vol,
      chap: String(chap),
      vol2: '',
      chap2: '',
      handwriting: '',
      curUrl: '/read/reading.asp'
    }).toString();

    const godpiaRes = await fetch('https://www.godpia.com/read/reading_body.asp', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Referer': 'https://www.godpia.com/read/reading.asp'
      },
      body: postBody
    });

    if (!godpiaRes.ok) {
      return json({ success: false, error: `갓피아 서버 응답 오류 (HTTP ${godpiaRes.status})` }, 502);
    }

    const verses = parseGodpiaHtml(await godpiaRes.text());
    if (verses.length === 0) {
      return json({ success: false, error: '해당 장의 성경 본문을 불러오지 못했습니다.' }, 404);
    }

    const response = json(
      {
        success: true,
        data: {
          vol,
          chapter: chap,
          version: { code: ver, name: BIBLE_VERSIONS[ver] },
          verses
        }
      },
      200,
      { 'Cache-Control': 'public, max-age=86400' }
    );

    if (cache) {
      const put = cache.put(cacheKey, response.clone());
      if (context.waitUntil) context.waitUntil(put);
      else await put;
    }

    return response;
  } catch (err) {
    return json({ success: false, error: `서버 내부 오류: ${err.message}` }, 500);
  }
}
