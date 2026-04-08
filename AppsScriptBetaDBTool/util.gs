function decodeDuckDuckGoRedirect_(url) {
  var text = String(url || '');
  var match = text.match(/[?&]uddg=([^&]+)/);
  if (match) return decodeURIComponent(match[1]);
  return text;
}

function extractChordCandidatesFromChordLikeTags_(html) {
  var matches = [];
  var regex = /<(span|div|p|li|td|th|rt)\b[^>]*(?:class|id)\s*=\s*["'][^"']*(?:chord|code|key|kcode|fret|rt)[^"']*["'][^>]*>([\s\S]*?)<\/\1>/gi;
  var match;
  while ((match = regex.exec(html)) && matches.length < 240) {
    matches = matches.concat(extractChordTokensFromText_(stripTags_(match[2])));
  }
  return dedupeChordCandidates_(matches);
}

function extractChordCandidatesFromHtml_(html) {
  var raw = decodeHtmlEntities_(String(html || ''));
  if (!raw) return [];
  var structured = extractChordCandidatesFromChordLikeTags_(raw);
  if (structured.length >= 8) return structured;
  var lineBased = extractChordCandidatesFromTextLines_(raw);
  if (lineBased.length >= 8) return lineBased;
  return dedupeChordCandidates_(extractChordTokensFromText_(stripTags_(raw)));
}

function stripTags_(html) {
  return String(html || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function isSupportedYoutubeValue_(rawValue) { return !!buildYoutubeEmbedUrl_(rawValue); }

function buildYoutubeSearchUrl_(artist, title) {
  var query = [artist || '', title || '', 'official'].join(' ').trim();
  return query ? 'https://www.youtube.com/results?search_query=' + encodeURIComponent(query) : '';
}

function decodeSearchRedirectUrl_(url) {
  var text = String(url || '');
  if (/duckduckgo\.com/.test(text) && /[?&]uddg=/.test(text)) return decodeDuckDuckGoRedirect_(text);
  if (/bing\.com\/ck\//.test(text)) {
    var targetMatch = text.match(/[?&]u=([^&]+)/);
    if (targetMatch) {
      try {
        var decoded = decodeURIComponent(targetMatch[1]);
        var cleaned = decoded.replace(/^a1/, '');
        var urlMatch = cleaned.match(/https?:\/\/.*/);
        if (urlMatch) return urlMatch[0];
      } catch (_error) {}
    }
  }
  return text;
}

function extractYoutubeVideoId_(url) {
  var text = String(url || '').trim();
  var match = text.match(/(?:v=|youtu\.be\/|embed\/|shorts\/)([A-Za-z0-9_-]{11})/);
  return match ? match[1] : '';
}

function sanitizeUrlInput_(value) { return String(value == null ? '' : value).trim(); }

function normalizePartKey_(partText) {
  var normalized = String(partText || '').trim();
  var map = { intro:'intro', 'イントロ':'intro', A:'A', 'Aメロ':'A', B:'B', 'Bメロ':'B', 'サビ':'サビ', chorus:'サビ' };
  return map.hasOwnProperty(normalized) ? map[normalized] : normalized;
}

function dedupeChordCandidates_(matches) {
  var blacklist = { HTML:true, HTTP:true, HTTPS:true, JPG:true, PNG:true, SVG:true };
  var filtered = [];
  var repeatGuard = 0;
  var prev = '';
  (matches || []).forEach(function(chord) {
    if (!chord || blacklist[chord]) return;
    if (chord === prev) {
      repeatGuard += 1;
      if (repeatGuard > 2) return;
    } else {
      repeatGuard = 0;
    }
    filtered.push(chord);
    prev = chord;
  });
  return filtered;
}

function extractChordCandidatesFromTextLines_(html) {
  var text = String(html || '')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<(br|\/p|\/div|\/li|\/tr|\/h\d)\b[^>]*>/gi, '\n');
  text = stripTags_(decodeHtmlEntities_(text));
  var lines = text.split(/\r?\n/);
  var matches = [];
  for (var i = 0; i < lines.length; i += 1) {
    var line = lines[i].trim();
    if (!line) continue;
    var tokens = extractChordTokensFromText_(line);
    if (tokens.length >= 2) matches = matches.concat(tokens);
    if (matches.length >= 240) break;
  }
  return dedupeChordCandidates_(matches);
}

function normalizeCandidateUrl_(url) {
  return String(url || '')
    .replace(/[),.;]+$/g, '')
    .replace(/&amp;/g, '&');
}

function buildSearchUrls_(query) {
  var encoded = encodeURIComponent(query);
  return [
    { name: 'duckduckgo', url: 'https://duckduckgo.com/html/?q=' + encoded },
    { name: 'bing', url: 'https://www.bing.com/search?q=' + encoded }
  ];
}

function normalizePartLabel_(partText) {
  var canonical = normalizePartKey_(partText);
  var map = { intro:'intro', A:'Aメロ', B:'Bメロ', 'サビ':'サビ' };
  return map[canonical] || canonical;
}

function extractSearchResultsFromHtml_(html) {
  var text = String(html || '');
  var results = [];
  var regex = /<a[^>]+class="[^"]*result__a[^"]*"[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>[\s\S]*?(?:<a[^>]+class="[^"]*result__snippet[^"]*"[^>]*>|<div[^>]+class="[^"]*result__snippet[^"]*"[^>]*>)([\s\S]*?)(?:<\/a>|<\/div>)/g;
  var match;
  while ((match = regex.exec(text)) && results.length < 12) {
    results.push({
      url: decodeHtmlEntities_(stripTags_(match[1])),
      title: decodeHtmlEntities_(stripTags_(match[2])),
      snippet: decodeHtmlEntities_(stripTags_(match[3]))
    });
  }
  if (!results.length) {
    var fallbackRegex = /<a[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/g;
    while ((match = fallbackRegex.exec(text)) && results.length < 12) {
      var url = decodeHtmlEntities_(stripTags_(match[1]));
      if (url.indexOf('uddg=') < 0) continue;
      results.push({ url: url, title: decodeHtmlEntities_(stripTags_(match[2])), snippet: '' });
    }
  }
  return results;
}

function normalizeBaseId_(baseId) {
  var text = normalizeNumericString_(baseId);
  if (!text) throw new Error('先頭IDは数値で入力してください。');
  return Number(text);
}

function fetchHtmlText_(url) {
  var response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (compatible; GoRockCampBot/1.0; +https://script.google.com/)'
    }
  });
  var status = response.getResponseCode();
  if (status < 200 || status >= 300) throw new Error('HTTP ' + status);
  return response.getContentText();
}

function normalizeNumericString_(value) { var text = String(value == null ? '' : value).trim(); return /^\d+$/.test(text) ? text : ''; }

function extractTargetLinksFromHtml_(html, domains) {
  var text = decodeHtmlEntities_(String(html || ''));
  var escaped = text.replace(/\\\//g, '/');
  var links = [];
  var regex = /(https?:\/\/[^\s"'<>\\)]+)/g;
  var match;
  while ((match = regex.exec(escaped)) && links.length < 120) {
    var url = match[1];
    if (!domains.some(function(domain) { return url.indexOf(domain) >= 0; })) continue;
    links.push(url);
  }
  return links;
}

function decodeHtmlEntities_(text) {
  return String(text || '')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
}

function toBooleanLikeString_(value) { var text = String(value == null ? '' : value).trim().toUpperCase(); return text === 'TRUE' || text === 'FALSE' ? text : (text || ''); }

function buildBarsFromChordPool_(chordPool, startIndex, barCount) {
  var source = (chordPool || []).filter(Boolean);
  var bars = [];
  if (!source.length) {
    for (var i = 0; i < barCount; i += 1) bars.push({ bar: i + 1, firstHalf: '', secondHalf: '' });
    return bars;
  }
  var cursor = Number(startIndex || 0);
  for (var bar = 0; bar < barCount; bar += 1) {
    var first = source[cursor % source.length] || '';
    var second = source[(cursor + 1) % source.length] || '';
    if (source.length === 1) second = '';
    bars.push({ bar: bar + 1, firstHalf: first, secondHalf: second });
    cursor += 2;
  }
  return bars;
}

function extractChordTokensFromText_(text) {
  var normalized = String(text || '').replace(/[|｜]/g, ' ');
  var regex = /\b[A-G](?:#|b)?(?:maj7|M7|m7-5|mM7|m7|m6|m|7|6|9|11|13|add9|sus2|sus4|dim|aug)?(?:\([b#]?(?:5|9|11|13)\))?(?:\/[A-G](?:#|b)?)?\b/g;
  return normalized.match(regex) || [];
}

function buildYoutubeEmbedUrl_(rawValue) {
  var text = String(rawValue || '').trim();
  if (!text) return '';
  var idMatch = text.match(/(?:v=|youtu\.be\/|embed\/|shorts\/)([A-Za-z0-9_-]{11})/);
  var videoId = idMatch ? idMatch[1] : text;
  return /^[A-Za-z0-9_-]{11}$/.test(videoId) ? 'https://www.youtube.com/embed/' + videoId : '';
}

