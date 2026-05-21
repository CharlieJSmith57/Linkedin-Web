// greek_merge.js
// Greek life org database, alias resolution, chapter parsing.
// Embedded inline in index.html — not a separate file.

// ── Org database (mirrors greek_orgs.json, kept in sync manually) ─────────────
const GREEK_DB = {
  fraternities: [
    { id:"pike",    full:"Pi Kappa Alpha",    letters:"ΠΚΑ", nick:"PIKE",      aliases:["Pike","PIKE","Pi Kappa Alpha","ΠΚΑ","PiKA","Pikes"],           p1:"#6B0D2E", p2:"#C8A800" },
    { id:"fiji",    full:"Phi Gamma Delta",   letters:"ΦΓΔ", nick:"FIJI",      aliases:["FIJI","Phi Gamma Delta","ΦΓΔ","Phi Gam","Phi Gamma Delta"],    p1:"#4B2E8C", p2:"#C8B000" },
    { id:"sigep",   full:"Sigma Phi Epsilon", letters:"ΣΦΕ", nick:"Sig Ep",    aliases:["Sig Ep","SigEp","Sigma Phi Epsilon","ΣΦΕ"],                    p1:"#7B0000", p2:"#4A1E8C" },
    { id:"sigchi",  full:"Sigma Chi",         letters:"ΣΧ",  nick:"Sig Chi",   aliases:["Sig Chi","Sigma Chi","ΣΧ","SigChi"],                           p1:"#003087", p2:"#C8B400" },
    { id:"sae",     full:"Sigma Alpha Epsilon",letters:"ΣΑΕ",nick:"SAE",       aliases:["SAE","Sigma Alpha Epsilon","ΣΑΕ"],                             p1:"#4B2E8C", p2:"#C8A800" },
    { id:"phidelt", full:"Phi Delta Theta",   letters:"ΦΔΘ", nick:"Phi Delt",  aliases:["Phi Delt","Phi Delta Theta","ΦΔΘ","Phi"],                     p1:"#003087", p2:"#87CEEB" },
    { id:"beta",    full:"Beta Theta Pi",     letters:"ΒΘΠ", nick:"Beta",      aliases:["Beta","Beta Theta Pi","ΒΘΠ","Beta Theta"],                    p1:"#CC0000", p2:"#FFC0CB" },
    { id:"kapsig",  full:"Kappa Sigma",       letters:"ΚΣ",  nick:"Kappa Sig", aliases:["Kappa Sig","Kappa Sigma","ΚΣ"],                               p1:"#003087", p2:"#CC0000" },
    { id:"lambdachi",full:"Lambda Chi Alpha", letters:"ΛΧΑ", nick:"Lambda Chi",aliases:["Lambda Chi","Lambda Chi Alpha","ΛΧΑ","Chops"],                p1:"#4B2E8C", p2:"#C8A800" },
    { id:"thetachi",full:"Theta Chi",         letters:"ΘΧ",  nick:"T Chi",     aliases:["Theta Chi","ΘΧ","T Chi","Ox","ThetaChi"],                    p1:"#CC0000", p2:"#F5F5DC" },
    { id:"delt",    full:"Delta Tau Delta",   letters:"ΔΤΔ", nick:"Delt",      aliases:["Delt","Delta Tau Delta","ΔΤΔ","DTD"],                         p1:"#4B2E8C", p2:"#C8A800" },
    { id:"sigmanu", full:"Sigma Nu",          letters:"ΣΝ",  nick:"Sig Nu",    aliases:["Sig Nu","Sigma Nu","ΣΝ","Knights"],                           p1:"#1C1C1C", p2:"#F5F5DC" },
    { id:"ato",     full:"Alpha Tau Omega",   letters:"ΑΤΩ", nick:"ATO",       aliases:["ATO","Alpha Tau Omega","ΑΤΩ","Tau"],                         p1:"003087",  p2:"#C8A800" },
    { id:"tke",     full:"Tau Kappa Epsilon", letters:"ΤΚΕ", nick:"TKE",       aliases:["TKE","Tau Kappa Epsilon","ΤΚΕ","Teke"],                      p1:"#CC0000", p2:"#808080" },
    { id:"phipsi",  full:"Phi Kappa Psi",     letters:"ΦΚΨ", nick:"Phi Psi",   aliases:["Phi Psi","Phi Kappa Psi","ΦΚΨ"],                             p1:"#006400", p2:"#CC0000" },
    { id:"phikaptau",full:"Phi Kappa Tau",    letters:"ΦΚΤ", nick:"Phi Tau",   aliases:["Phi Tau","Phi Kappa Tau","ΦΚΤ"],                             p1:"#CC0000", p2:"#C8A800" },
    { id:"dke",     full:"Delta Kappa Epsilon",letters:"ΔΚΕ",nick:"Deke",      aliases:["Deke","DKE","Delta Kappa Epsilon","ΔΚΕ"],                    p1:"#CC0000", p2:"#003087" },
    { id:"pikapp",  full:"Pi Kappa Phi",      letters:"ΠΚΦ", nick:"Pi Kapp",   aliases:["Pi Kapp","Pi Kappa Phi","ΠΚΦ","PKP"],                       p1:"#C8A800", p2:"#003087" },
    { id:"sigmapi", full:"Sigma Pi",          letters:"ΣΠ",  nick:"Sig Pi",    aliases:["Sig Pi","Sigma Pi","ΣΠ"],                                    p1:"#C8A4D0", p2:"#F5F5DC" },
    { id:"aepi",    full:"Alpha Epsilon Pi",  letters:"ΑΕΠ", nick:"AEPi",      aliases:["AEPi","Alpha Epsilon Pi","ΑΕΠ","Ape"],                       p1:"003087",  p2:"#C8A800" },
    { id:"alphasig",full:"Alpha Sigma Phi",   letters:"ΑΣΦ", nick:"Alpha Sig", aliases:["Alpha Sig","Alpha Sigma Phi","ΑΣΦ"],                         p1:"#CC0000", p2:"#C8A800" },
    { id:"apa",     full:"Alpha Phi Alpha",   letters:"ΑΦΑ", nick:"Alpha",     aliases:["Alpha Phi Alpha","ΑΦΑ","APhiA"],                             p1:"#1C1C1C", p2:"#C8A800" },
    { id:"kapsi",   full:"Kappa Alpha Psi",   letters:"ΚΑΨ", nick:"Kappa",     aliases:["Kappa Alpha Psi","ΚΑΨ"],                                    p1:"#CC0000", p2:"#F5F5DC" },
    { id:"omegaps", full:"Omega Psi Phi",     letters:"ΩΨΦ", nick:"Ques",      aliases:["Omega Psi Phi","ΩΨΦ","Ques"],                               p1:"#4B2E8C", p2:"#C8A800" },
    { id:"phibets", full:"Phi Beta Sigma",    letters:"ΦΒΣ", nick:"Sigmas",    aliases:["Phi Beta Sigma","ΦΒΣ","Sigmas"],                            p1:"#003087", p2:"#F5F5DC" }
  ],
  sororities: [
    { id:"adpi",    full:"Alpha Delta Pi",    letters:"ΑΔΠ", nick:"ADPi",      aliases:["ADPi","Alpha Delta Pi","ΑΔΠ","A D Pi"],                     p1:"#003087", p2:"#006400" },
    { id:"aephi",   full:"Alpha Epsilon Phi", letters:"ΑΕΦ", nick:"AEPhi",     aliases:["AEPhi","Alpha Epsilon Phi","ΑΕΦ","A E Phi"],                p1:"#006400", p2:"#C8A800" },
    { id:"aphi",    full:"Alpha Phi",         letters:"ΑΦ",  nick:"A Phi",     aliases:["Alpha Phi","ΑΦ","A Phi","APhi"],                            p1:"#800020", p2:"#C0C0C0" },
    { id:"axo",     full:"Alpha Chi Omega",   letters:"ΑΧΩ", nick:"Alpha Chi", aliases:["Alpha Chi","Alpha Chi Omega","ΑΧΩ","AXO"],                  p1:"#CC0000", p2:"#006400" },
    { id:"dg",      full:"Delta Gamma",       letters:"ΔΓ",  nick:"DG",        aliases:["DG","Delta Gamma","ΔΓ","D G"],                              p1:"#B87333", p2:"#FFC0CB" },
    { id:"tridelt", full:"Delta Delta Delta", letters:"ΔΔΔ", nick:"Tri Delta", aliases:["Tri Delta","Tri Delt","Delta Delta Delta","ΔΔΔ","DDD"],     p1:"#C0C0C0", p2:"#C8A800" },
    { id:"dz",      full:"Delta Zeta",        letters:"ΔΖ",  nick:"DZ",        aliases:["DZ","Delta Zeta","ΔΖ","D Z"],                              p1:"#CC0000", p2:"#006400" },
    { id:"gammaphi",full:"Gamma Phi Beta",    letters:"ΓΦΒ", nick:"Gamma Phi", aliases:["Gamma Phi","Gamma Phi Beta","ΓΦΒ"],                         p1:"#4B2E8C", p2:"#C0BF94" },
    { id:"kat",     full:"Kappa Alpha Theta", letters:"ΚΑΘ", nick:"Theta",     aliases:["Theta","Kappa Alpha Theta","ΚΑΘ","KAT"],                   p1:"#1C1C1C", p2:"#C8A800" },
    { id:"kd",      full:"Kappa Delta",       letters:"ΚΔ",  nick:"KD",        aliases:["KD","Kappa Delta","ΚΔ","K D"],                             p1:"#006400", p2:"#F5F5DC" },
    { id:"kkg",     full:"Kappa Kappa Gamma", letters:"ΚΚΓ", nick:"Kappa",     aliases:["Kappa","KKG","Kappa Kappa Gamma","ΚΚΓ","K K G"],           p1:"#003087", p2:"#87CEFA" },
    { id:"piphi",   full:"Pi Beta Phi",       letters:"ΠΒΦ", nick:"Pi Phi",    aliases:["Pi Phi","Pi Beta Phi","ΠΒΦ","PBP"],                        p1:"#800020", p2:"#87CEFA" },
    { id:"zta",     full:"Zeta Tau Alpha",    letters:"ΖΤΑ", nick:"Zeta",      aliases:["Zeta","ZTA","Zeta Tau Alpha","ΖΤΑ"],                       p1:"#008B8B", p2:"#808080" },
    { id:"sigmakap",full:"Sigma Kappa",       letters:"ΣΚ",  nick:"Sig Kap",   aliases:["Sig Kap","Sigma Kappa","ΣΚ","SK"],                         p1:"#4B2E8C", p2:"#800020" },
    { id:"chiomega",full:"Chi Omega",         letters:"ΧΩ",  nick:"Chi O",     aliases:["Chi O","Chi Omega","ΧΩ","XO"],                             p1:"#CC0000", p2:"#C8A800" },
    { id:"dst",     full:"Delta Sigma Theta", letters:"ΔΣΘ", nick:"Deltas",    aliases:["Delta Sigma Theta","ΔΣΘ","Deltas","DST"],                 p1:"#CC0000", p2:"#F5F5DC" }
  ]
};

// Flat lookup: any alias/nick/letters/full → org entry
const GREEK_ALIAS_MAP = {};
[...GREEK_DB.fraternities, ...GREEK_DB.sororities].forEach(org => {
  org.aliases.forEach(a => {
    GREEK_ALIAS_MAP[a.toLowerCase().trim()] = org;
  });
});

/**
 * Resolve any string to a canonical org entry, or null.
 */
function resolveOrg(name) {
  if (!name) return null;
  return GREEK_ALIAS_MAP[name.toLowerCase().trim()] || null;
}

/**
 * Given a raw greek_orgs array entry like "PIKE, Beta Chapter at UT Austin"
 * parse into { org, chapter_name, school }.
 * Also handles bare entries like "Pi Kappa Alpha" or "FIJI".
 */
function parseGreekEntry(raw) {
  if (!raw) return null;
  // Try: "LETTERS, Chapter at School" or "Full Name, Chapter"
  const commaIdx = raw.indexOf(',');
  let orgPart = raw, chapterPart = '';
  if (commaIdx > 0) {
    orgPart    = raw.slice(0, commaIdx).trim();
    chapterPart = raw.slice(commaIdx + 1).trim();
  }
  const org = resolveOrg(orgPart) || resolveOrg(raw.trim());
  if (!org) return null;

  // Parse chapter from chapterPart: "Beta Chapter at UT Austin" → {name:"Beta Chapter", school:"UT Austin"}
  let chapterName = '', school = '';
  if (chapterPart) {
    const atIdx = chapterPart.toLowerCase().indexOf(' at ');
    if (atIdx >= 0) {
      chapterName = chapterPart.slice(0, atIdx).trim();
      school      = chapterPart.slice(atIdx + 4).trim();
    } else {
      chapterName = chapterPart;
    }
  }

  return { org, chapter_name: chapterName, school };
}

/**
 * Get all parsed greek entries for a connection.
 * Returns array of {org, chapter_name, school}.
 */
function getGreekEntries(c) {
  return (c.greek_orgs || [])
    .map(parseGreekEntry)
    .filter(Boolean);
}

/**
 * User-editable greek merges — stored in localStorage alongside company merges.
 * Format: { "alias": "canonical_org_id" }
 */
function getGreekMerges() {
  const o = getOverrides(); // from main app
  return o.greek_merges || {};
}
function saveGreekMerge(alias, orgId) {
  const o = getOverrides();
  o.greek_merges = o.greek_merges || {};
  o.greek_merges[alias.toLowerCase().trim()] = orgId;
  saveOverrides(o);
  // Re-add to alias map
  const org = [...GREEK_DB.fraternities,...GREEK_DB.sororities].find(x=>x.id===orgId);
  if (org) GREEK_ALIAS_MAP[alias.toLowerCase().trim()] = org;
}
function removeGreekMerge(alias) {
  const o = getOverrides();
  if (o.greek_merges) delete o.greek_merges[alias.toLowerCase().trim()];
  saveOverrides(o);
  delete GREEK_ALIAS_MAP[alias.toLowerCase().trim()];
}

// Apply saved greek merges on load
function applyGreekMerges() {
  const merges = getGreekMerges();
  Object.entries(merges).forEach(([alias, orgId]) => {
    const org = [...GREEK_DB.fraternities,...GREEK_DB.sororities].find(x=>x.id===orgId);
    if (org) GREEK_ALIAS_MAP[alias] = org;
  });
}
