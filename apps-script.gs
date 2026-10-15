/**
 * Tryllepub – påmeldinger
 * Lim inn i regnearket via Utvidelser → Apps Script, og publiser som nettapp.
 * Tar imot påmeldinger fra påmeldingssiden, lagrer dem i arket
 * og svarer med hvor mange plasser som er igjen.
 */

const KAPASITET = 50;
const ARK = 'Påmeldinger';
const KOLONNER = ['Tidspunkt', 'Forestilling', 'Navn', 'E-post', 'Mobil',
  'Ordinære', 'Student', 'Totalt antall', 'Beløp (kr)', 'Nyhetsbrev', 'Betalt'];
const PRIS_ORDINAER = 200;
const PRIS_STUDENT = 100;

function hentArk_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ark = ss.getSheetByName(ARK);
  if (!ark) {
    ark = ss.insertSheet(ARK);
    ark.appendRow(KOLONNER);
    ark.setFrozenRows(1);
    ark.getRange(1, 1, 1, KOLONNER.length).setFontWeight('bold');
  }
  return ark;
}

function solgt_(ark, forestilling) {
  const rader = ark.getDataRange().getValues().slice(1);
  return rader
    .filter(r => String(r[1]) === forestilling)
    .reduce((sum, r) => sum + (Number(r[7]) || 0), 0);
}

function svar_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Nedtelling: ?forestilling=2026-11-11
function doGet(e) {
  const forestilling = String((e.parameter && e.parameter.forestilling) || '');
  const ark = hentArk_();
  const igjen = Math.max(0, KAPASITET - solgt_(ark, forestilling));
  return svar_({ ok: true, igjen: igjen, kapasitet: KAPASITET });
}

// Påmelding
function doPost(e) {
  const lås = LockService.getScriptLock();
  lås.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const forestilling = String(d.forestilling || '').slice(0, 20);
    const navn = String(d.navn || '').trim().slice(0, 100);
    const epost = String(d.epost || '').trim().slice(0, 150);
    const mobil = String(d.mobil || '').trim().slice(0, 30);
    const ordinaer = Math.max(0, Math.min(10, parseInt(d.ordinaer, 10) || 0));
    const student = Math.max(0, Math.min(10, parseInt(d.student, 10) || 0));
    const totalt = ordinaer + student;

    if (!forestilling || !navn || (!epost && !mobil) || totalt < 1) {
      return svar_({ ok: false, feil: 'Mangler navn, kontaktinfo eller antall billetter.' });
    }

    const ark = hentArk_();
    const igjen = KAPASITET - solgt_(ark, forestilling);
    if (totalt > igjen) {
      return svar_({ ok: false, feil: igjen > 0
        ? 'Det er bare ' + igjen + ' plasser igjen.'
        : 'Forestillingen er dessverre fullt booket.', igjen: Math.max(0, igjen) });
    }

    const belop = ordinaer * PRIS_ORDINAER + student * PRIS_STUDENT;
    ark.appendRow([new Date(), forestilling, navn, epost, "'" + mobil, ordinaer, student,
      totalt, belop, d.nyhetsbrev ? 'Ja' : 'Nei', '']);

    return svar_({ ok: true, igjen: igjen - totalt, belop: belop });
  } catch (err) {
    return svar_({ ok: false, feil: 'Noe gikk galt. Prøv igjen.' });
  } finally {
    lås.releaseLock();
  }
}
