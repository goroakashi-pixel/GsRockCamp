function inferOriginalKeyFromBundle_(bundle) {
  var chords = [];
  bundle.rows.forEach(function(row) {
    row.originalChords.forEach(function(cell) {
      splitChordCell_(cell).aiSuggestion.forEach(function(chord) { chords.push(chord); });
    });
  });
  return inferOriginalKeyFromChordList_(chords);
}

function getScaleSetForKey_(keyName) {
  var normalized = normalizeKeyName_(keyName);
  if (/m$/.test(normalized) && normalized !== 'Am') return getNaturalMinorScaleSet_(normalized.replace(/m$/, ''));
  if (normalized === 'Am') return getNaturalMinorScaleSet_('A');
  return getMajorScaleSet_(normalized || 'C');
}

function buildNonDiatonicCell_(cellText, keyName) {
  var parts = splitChordCell_(cellText).aiSuggestion;
  if (!parts.length) return '';
  return parts.some(function(chord) { return !isChordDiatonic_(chord, keyName); }) ? 'TRUE' : 'FALSE';
}

function sanitizeChordInput_(value) { return String(value == null ? '' : value).replace(/\s+/g, ' ').trim().replace(/｜/g, '│'); }

function getDegreeText_(chordText, keyName) {
  var parsed = parseChordSymbol_(chordText);
  return getRomanDegree_(parsed.root, keyName) + inferDegreeSuffix_(parsed.quality);
}

function determineQuizKey_(originalKey) {
  var normalized = sanitizeKeyInput_(originalKey);
  return normalized && /m$/.test(normalized) ? 'Am' : 'C';
}

function getFunctionText_(chordText, keyName) {
  var parsed = parseChordSymbol_(chordText);
  var roman = getRomanDegree_(parsed.root, keyName);
  if (/Ⅴ|Ⅶ/.test(roman)) return isChordDiatonic_(chordText, keyName) ? 'D' : 'セカンダリードミナント';
  if (/Ⅱ|Ⅳ/.test(roman)) return isChordDiatonic_(chordText, keyName) ? 'SD' : '同主調借用';
  if (/Ⅰ|Ⅵ|Ⅲ/.test(roman)) return isChordDiatonic_(chordText, keyName) ? 'T' : 'モーダルインターチェンジ';
  return 'クロマチック';
}

function recalcTheory_(originalChords, originalKey, quizKey) {
  var sourceKey = normalizeKeyName_(originalKey || quizKey || 'C');
  var targetKey = normalizeKeyName_(quizKey || originalKey || 'C');
  var shift = getSemitoneShift_(sourceKey, targetKey);
  return {
    changeChord: originalChords.map(function(cell) { return transposeCell_(cell, shift); }),
    degree: originalChords.map(function(cell) { return buildTheoryCell_(cell, function(chord) { return getDegreeText_(chord, targetKey); }); }),
    functionText: originalChords.map(function(cell) { return buildTheoryCell_(cell, function(chord) { return getFunctionText_(chord, targetKey); }); }),
    nonDiatonic: originalChords.map(function(cell) { return buildNonDiatonicCell_(cell, targetKey); })
  };
}

function transposePitchName_(noteName, shift) {
  var index = noteNameToIndex_(noteName);
  return index < 0 ? noteName : indexToSharpName_((index + shift) % 12);
}

function transposeCell_(cellText, shift) {
  var parts = splitChordCell_(cellText).aiSuggestion;
  return parts.length ? parts.map(function(chord) { return transposeChordSymbol_(chord, shift); }).join('│') : '';
}

function inferOriginalKeyFromChordList_(chords) {
  if (!chords || !chords.length) return '';
  var candidates = ['C','G','D','A','E','F','Bb','Eb','Am','Em','Dm','Bm','Cm','Fm'];
  var bestKey = '';
  var bestScore = -1;
  candidates.forEach(function(candidate) {
    var score = 0;
    chords.forEach(function(chord, index) {
      var parsed = parseChordSymbol_(chord);
      var roman = getRomanDegree_(parsed.root, candidate);
      if (isChordDiatonic_(chord, candidate)) score += 3;
      if (/Ⅰ|Ⅳ|Ⅴ|Ⅵ/.test(roman)) score += 2;
      if (index === 0 && /^(Ⅰ|Ⅵ)/.test(roman)) score += 1;
      if (index === chords.length - 1 && /^(Ⅰ|Ⅵ)/.test(roman)) score += 2;
    });
    if (score > bestScore) { bestScore = score; bestKey = candidate; }
  });
  return bestKey;
}

function validateChordSymbol_(chordText, part, barNumber) {
  var slashCount = (String(chordText).match(/\//g) || []).length;
  if (slashCount > 1) throw new Error(part + ' ' + barNumber + '小節 に不正な / 連続表記があります: ' + chordText);
  var base = String(chordText).split('/')[0];
  if (!/^([A-G](?:#|b|♭)?)(?:[0-9A-Za-z+()#♭b\-susdimaugmajMadd]*)$/.test(base)) throw new Error(part + ' ' + barNumber + '小節 のコード形式が不正です: ' + chordText);
  if (slashCount === 1) {
    var bass = String(chordText).split('/')[1] || '';
    if (!/^[A-G](?:#|b|♭)?$/.test(bass)) throw new Error(part + ' ' + barNumber + '小節 のオンコード表記が不正です: ' + chordText);
  }
}

function getSemitoneShift_(fromKey, toKey) {
  return ((noteNameToIndex_(normalizeKeyRoot_(toKey)) - noteNameToIndex_(normalizeKeyRoot_(fromKey))) % 12 + 12) % 12;
}

function getMajorScaleSet_(keyName) {
  var root = noteNameToIndex_(keyName);
  var set = {};
  [0, 2, 4, 5, 7, 9, 11].forEach(function(interval) { set[indexToSharpName_((root + interval) % 12)] = true; });
  return set;
}

function indexToSharpName_(index) { return ['C','C#','D','Eb','E','F','F#','G','Ab','A','Bb','B'][((index % 12) + 12) % 12]; }

function noteNameToIndex_(noteName) {
  var normalized = normalizeKeyName_(noteName);
  var map = {C:0,'B#':0,'C#':1,Db:1,D:2,'D#':3,Eb:3,E:4,Fb:4,'E#':5,F:5,'F#':6,Gb:6,G:7,'G#':8,Ab:8,A:9,'A#':10,Bb:10,B:11,Cb:11};
  return map.hasOwnProperty(normalized) ? map[normalized] : -1;
}

function normalizeKeyName_(text) { return String(text || '').trim().replace(/♭/g, 'b').replace(/＃/g, '#').replace(/([A-Ga-g])/, function(match){ return match.toUpperCase(); }); }

function buildTheoryCell_(cellText, mapper) {
  var parts = splitChordCell_(cellText).aiSuggestion;
  return parts.length ? parts.map(mapper).join('│') : '';
}

function buildChordPitchClasses_(chordText) {
  var parsed = parseChordSymbol_(chordText);
  var rootIndex = noteNameToIndex_(parsed.root);
  if (rootIndex < 0) return [];
  var intervals = [0, 4, 7];
  if (/^m(?!aj)/.test(parsed.quality)) intervals = [0, 3, 7];
  if (/dim/.test(parsed.quality)) intervals = [0, 3, 6];
  if (/aug/.test(parsed.quality)) intervals = [0, 4, 8];
  if (/sus4/.test(parsed.quality)) intervals = [0, 5, 7];
  if (/6/.test(parsed.quality)) intervals.push(9);
  if (/maj7|M7|mM7/.test(parsed.quality)) intervals.push(11); else if (/7/.test(parsed.quality)) intervals.push(10);
  if (/9|add9/.test(parsed.quality)) intervals.push(14);
  if (parsed.bass) intervals.push(noteNameToIndex_(parsed.bass) - rootIndex);

  var pitchSet = {};
  intervals.forEach(function(interval) { pitchSet[indexToSharpName_((rootIndex + interval) % 12)] = true; });
  return Object.keys(pitchSet);
}

function splitChordCell_(cellText) {
  var text = String(cellText || '').trim();
  if (!text) return { first: '', second: '', aiSuggestion: [] };
  var parts = text.split('│').map(function(item) { return item.trim(); }).filter(Boolean);
  return { first: parts[0] || '', second: parts[1] || '', aiSuggestion: parts };
}

function getNaturalMinorScaleSet_(keyName) {
  var root = noteNameToIndex_(keyName);
  var set = {};
  [0, 2, 3, 5, 7, 8, 10].forEach(function(interval) { set[indexToSharpName_((root + interval) % 12)] = true; });
  return set;
}

function sanitizeKeyInput_(value) { return normalizeKeyName_(value).replace(/major$/i, '').replace(/minor$/i, 'm'); }

function transposeChordSymbol_(chordText, shift) {
  var parsed = parseChordSymbol_(chordText);
  var root = transposePitchName_(parsed.root, shift);
  var bass = parsed.bass ? transposePitchName_(parsed.bass, shift) : '';
  return root + parsed.quality + (bass ? '/' + bass : '');
}

function getRomanDegree_(rootName, keyName) {
  var tonic = normalizeKeyRoot_(keyName || 'C');
  var diff = ((noteNameToIndex_(rootName) - noteNameToIndex_(tonic)) % 12 + 12) % 12;
  var map = {0:'Ⅰ',1:'bⅡ',2:'Ⅱ',3:'bⅢ',4:'Ⅲ',5:'Ⅳ',6:'bⅤ',7:'Ⅴ',8:'bⅥ',9:'Ⅵ',10:'bⅦ',11:'Ⅶ'};
  return map[diff] || '?';
}

function normalizeKeyRoot_(keyName) {
  var normalized = normalizeKeyName_(keyName || 'C');
  if (normalized === 'Am') return 'A';
  return normalized.replace(/m$/, '');
}

function inferDegreeSuffix_(quality) {
  quality = String(quality || '');
  if (!quality) return '';
  if (/m7-5/.test(quality)) return 'm7-5';
  if (/maj7|M7/.test(quality)) return 'M7';
  if (/m6/.test(quality)) return 'm6';
  if (/6/.test(quality)) return '6';
  if (/m7/.test(quality)) return 'm7';
  if (/m(?!aj)/.test(quality)) return 'm';
  if (/dim/.test(quality)) return 'dim';
  if (/aug/.test(quality)) return 'aug';
  if (/sus4/.test(quality)) return 'sus4';
  if (/add9/.test(quality)) return '(add9)';
  if (/9/.test(quality)) return '9';
  if (/7/.test(quality)) return '7';
  return quality;
}

function isChordDiatonic_(chordText, keyName) {
  var scaleSet = getScaleSetForKey_(keyName);
  return buildChordPitchClasses_(chordText).every(function(note) { return scaleSet[note]; });
}

function parseChordSymbol_(chordText) {
  var halves = String(chordText || '').trim().split('/');
  var main = halves[0] || '';
  var bass = halves[1] || '';
  var match = main.match(/^([A-G](?:#|b|♭)?)(.*)$/);
  if (!match) throw new Error('コードを解釈できません: ' + chordText);
  return { root: normalizeKeyName_(match[1]), quality: match[2] || '', bass: bass ? normalizeKeyName_(bass) : '' };
}

