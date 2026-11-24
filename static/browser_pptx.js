/* Pure-JS PPTX engine (requires JSZip) for Caminata Gemba.
 * Works identically in any modern browser (file:// or https://) and Node.js.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.GembaPPTX = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const SLIDE_W = 12192000;
  const SLIDE_H = 6858000;
  const COVER_BOX = [152400, 757937, 11887200, 5342117];
  const ANTES_BOX = [152400, 1603275, 5784600, 4901025];
  const DESPUES_BOX = [6255000, 1603275, 5784600, 4901025];

  const REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
  const REL_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
  const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  const CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

  function escapeXml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function decodeXmlEntities(str) {
    return String(str)
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
  }

  function parseAttrs(tagStr) {
    const attrs = {};
    const re = /([\w:-]+)="([^"]*)"/g;
    let m;
    while ((m = re.exec(tagStr)) !== null) {
      attrs[m[1]] = m[2];
    }
    return attrs;
  }

  function parseRels(xml) {
    const rels = [];
    const re = /<Relationship\b[^>]*\/?>/g;
    let m;
    while ((m = re.exec(xml)) !== null) {
      const a = parseAttrs(m[0]);
      if (a.Id) {
        rels.push({ id: a.Id, type: a.Type || "", target: a.Target || "", raw: m[0] });
      }
    }
    return rels;
  }

  function extractTexts(xml) {
    const texts = [];
    const re = /<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g;
    let m;
    while ((m = re.exec(xml)) !== null) {
      const t = decodeXmlEntities(m[1]).trim();
      if (t) texts.push(t);
    }
    return texts;
  }

  function extractPictures(xml) {
    const pics = [];
    const re = /<p:pic\b[\s\S]*?<\/p:pic>/g;
    let m;
    while ((m = re.exec(xml)) !== null) {
      const block = m[0];
      const embedMatch = /r:embed="([^"]+)"/.exec(block);
      const offMatch = /<a:off\b[^>]*>/.exec(block);
      const extMatch = /<a:ext\b[^>]*>/.exec(block);
      const off = offMatch ? parseAttrs(offMatch[0]) : {};
      const ext = extMatch ? parseAttrs(extMatch[0]) : {};
      pics.push({
        rId: embedMatch ? embedMatch[1] : null,
        x: Number(off.x || 0),
        y: Number(off.y || 0),
        cx: Number(ext.cx || 0),
        cy: Number(ext.cy || 0),
        raw: block,
      });
    }
    return pics;
  }

  function fitBox(imgW, imgH, box) {
    const [bx, by, bcx, bcy] = box;
    const imgRatio = imgW / imgH;
    const boxRatio = bcx / bcy;
    let newW, newH;
    if (imgRatio > boxRatio) {
      newW = bcx;
      newH = Math.round(bcx / imgRatio);
    } else {
      newH = bcy;
      newW = Math.round(bcy * imgRatio);
    }
    const newX = bx + Math.round((bcx - newW) / 2);
    const newY = by + Math.round((bcy - newH) / 2);
    return [newX, newY, newW, newH];
  }

  function resolveMediaPath(target) {
    if (!target) return null;
    const clean = target.replace(/^(\.\.\/)+/, "").replace(/^\//, "");
    return clean.startsWith("ppt/") ? clean : `ppt/${clean}`;
  }

  function buildPicXml(shapeId, name, rId, x, y, cx, cy) {
    return (
      `<p:pic>` +
      `<p:nvPicPr><p:cNvPr id="${shapeId}" name="${escapeXml(name)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
      `</p:pic>`
    );
  }

  function buildNoteXml(shapeId, noteText, hasPhoto) {
    const [bx, by, bcx, bcy] = DESPUES_BOX;
    const x = hasPhoto ? bx + 100000 : bx + 250000;
    const y = hasPhoto ? by + bcy - 540000 : by + Math.round(bcy / 2) - 550000;
    const cx = hasPhoto ? bcx - 200000 : bcx - 500000;
    const cy = hasPhoto ? 480000 : 1100000;
    const sz = hasPhoto ? 2200 : 2800;
    return (
      `<p:sp>` +
      `<p:nvSpPr><p:cNvPr id="${shapeId}" name="GembaNote ${shapeId}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:srgbClr val="F1F5F9"/></a:solidFill>` +
      `<a:ln w="19050"><a:solidFill><a:srgbClr val="058DC7"/></a:solidFill></a:ln></p:spPr>` +
      `<p:txBody><a:bodyPr anchor="ctr" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"/>` +
      `<a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="es-ES" sz="${sz}" b="1">` +
      `<a:solidFill><a:srgbClr val="0F172A"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr>` +
      `<a:t>${escapeXml(noteText)}</a:t></a:r></a:p></p:txBody></p:sp>`
    );
  }

  function buildCoverSlideXml(rId, x, y, cx, cy) {
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
      `<p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
      buildPicXml(2, "Portada Gemba", rId, x, y, cx, cy) +
      `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
    );
  }

  function buildAntesSlideXml(rId, x, y, cx, cy) {
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
      `<p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="TopBar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="196952"/><a:ext cx="${SLIDE_W}" cy="1254000"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="058DC7"/></a:solidFill>` +
      `<a:ln w="12700"><a:solidFill><a:srgbClr val="1C3052"/></a:solidFill></a:ln></p:spPr></p:sp>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="AntesTitle"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="152375" y="362250"/><a:ext cx="5373300" cy="923400"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>` +
      `<a:ln w="76200"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></p:spPr>` +
      `<p:txBody><a:bodyPr anchor="ctr" wrap="square"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>` +
      `<a:r><a:rPr lang="es-ES" sz="5400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>ANTES</a:t></a:r>` +
      `</a:p></p:txBody></p:sp>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="4" name="DespuesTitle"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="5643241" y="362250"/><a:ext cx="6246000" cy="923400"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>` +
      `<a:ln w="76200"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></p:spPr>` +
      `<p:txBody><a:bodyPr anchor="ctr" wrap="square"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>` +
      `<a:r><a:rPr lang="es-ES" sz="5400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>DESPUÉS</a:t></a:r>` +
      `</a:p></p:txBody></p:sp>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="5" name="BottomBar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="6656697"/><a:ext cx="${SLIDE_W}" cy="187500"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="058DC7"/></a:solidFill>` +
      `<a:ln w="12700"><a:solidFill><a:srgbClr val="1C3052"/></a:solidFill></a:ln></p:spPr></p:sp>` +
      buildPicXml(6, "Foto ANTES", rId, x, y, cx, cy) +
      `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
    );
  }

  function buildSlideRelsXml(mediaFilename) {
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL_IMAGE}" Target="../media/${mediaFilename}"/>` +
      `</Relationships>`
    );
  }

  async function getOrderedSlides(zip) {
    const presXml = await zip.file("ppt/presentation.xml").async("string");
    const presRelsXml = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
    const relMap = {};
    for (const r of parseRels(presRelsXml)) {
      relMap[r.id] = r.target;
    }
    const ordered = [];
    const sldRe = /<p:sldId\b[^>]*\/?>/g;
    let m;
    while ((m = sldRe.exec(presXml)) !== null) {
      const a = parseAttrs(m[0]);
      const rId = a["r:id"];
      const target = relMap[rId];
      if (target) {
        const slidePath = target.startsWith("ppt/") ? target : `ppt/${target.replace(/^\//, "")}`;
        const relsPath = slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
        ordered.push({
          sldId: Number(a.id),
          rId,
          target,
          slidePath,
          relsPath,
          rawTag: m[0],
        });
      }
    }
    return { presXml, presRelsXml, ordered };
  }

  async function inspectZip(zip) {
    const { ordered } = await getOrderedSlides(zip);
    const walks = [];
    let currentWalk = null;
    let emptyTrailingCount = 0;

    for (let idx = 0; idx < ordered.length; idx++) {
      const entry = ordered[idx];
      const slideFile = zip.file(entry.slidePath);
      if (!slideFile) continue;
      const xml = await slideFile.async("string");
      const relsFile = zip.file(entry.relsPath);
      const relsXml = relsFile ? await relsFile.async("string") : "";
      const relMap = {};
      for (const r of parseRels(relsXml)) {
        relMap[r.id] = r.target;
      }

      const texts = extractTexts(xml);
      const pics = extractPictures(xml);
      const hasHeaders = texts.some((t) => t === "ANTES" || t === "DESPUÉS");

      // Cover slide: picture(s) without ANTES/DESPUÉS headers
      if (!hasHeaders && pics.length >= 1) {
        const coverMedia = resolveMediaPath(relMap[pics[0].rId]);
        currentWalk = {
          walk_index: walks.length + 1,
          cover_slide_index: idx,
          cover_slide_number: idx + 1,
          cover_media_path: coverMedia,
          slides: [],
        };
        walks.push(currentWalk);
        continue;
      }

      // Empty template slide
      const uniqueTexts = new Set(texts);
      if (pics.length === 0 && uniqueTexts.size === 2 && uniqueTexts.has("ANTES") && uniqueTexts.has("DESPUÉS")) {
        emptyTrailingCount++;
        continue;
      }

      if (!currentWalk) {
        currentWalk = {
          walk_index: 1,
          cover_slide_index: null,
          cover_slide_number: null,
          cover_media_path: null,
          slides: [],
        };
        walks.push(currentWalk);
      }

      const halfW = SLIDE_W / 2;
      let antesMedia = null;
      let despuesMedia = null;
      for (const pic of pics) {
        const centerX = pic.x + pic.cx / 2;
        const mPath = resolveMediaPath(relMap[pic.rId]);
        if (centerX < halfW) {
          if (!antesMedia) antesMedia = mPath;
        } else {
          if (!despuesMedia) despuesMedia = mPath;
        }
      }

      const extraNotes = texts.filter((t) => t !== "ANTES" && t !== "DESPUÉS").join(" ");
      currentWalk.slides.push({
        slide_index: idx,
        slide_number: idx + 1,
        slide_path: entry.slidePath,
        rels_path: entry.relsPath,
        has_antes: Boolean(antesMedia),
        has_despues: Boolean(despuesMedia || extraNotes),
        antes_media_path: antesMedia,
        despues_media_path: despuesMedia,
        notes: extraNotes,
      });
    }

    return {
      total_slides: ordered.length,
      empty_template_slides: emptyTrailingCount,
      walks,
    };
  }

  function findNextNumbers(zip, presXml, presRelsXml) {
    let maxSlideNum = 0;
    let maxMediaNum = 0;
    zip.forEach((relPath) => {
      const sm = /^ppt\/slides\/slide(\d+)\.xml$/.exec(relPath);
      if (sm) maxSlideNum = Math.max(maxSlideNum, Number(sm[1]));
      const mm = /^ppt\/media\/image(\d+)\./.exec(relPath);
      if (mm) maxMediaNum = Math.max(maxMediaNum, Number(mm[1]));
    });

    let maxSldId = 255;
    const idRe = /<p:sldId\b[^>]*\bid="(\d+)"/g;
    let m;
    while ((m = idRe.exec(presXml)) !== null) {
      maxSldId = Math.max(maxSldId, Number(m[1]));
    }

    let maxRId = 0;
    const rRe = /\bId="rId(\d+)"/g;
    while ((m = rRe.exec(presRelsXml)) !== null) {
      maxRId = Math.max(maxRId, Number(m[1]));
    }

    return { maxSlideNum, maxMediaNum, maxSldId, maxRId };
  }

  function ensureContentTypeDefault(ctXml, ext, contentType) {
    const re = new RegExp(`<Default\\b[^>]*Extension="${ext}"`, "i");
    if (re.test(ctXml)) return ctXml;
    return ctXml.replace(
      /(<Types\b[^>]*>)/,
      `$1<Default Extension="${ext}" ContentType="${contentType}"/>`
    );
  }

  /**
   * Append a cover slide + 1 ANTES/DESPUÉS slide per photo.
   * cover: { bytes: Uint8Array, width: number, height: number }
   * photos: Array<{ bytes: Uint8Array, width: number, height: number }>
   */
  async function appendWalk(zip, cover, photos, cleanTrailingEmpty = true) {
    let { presXml, presRelsXml, ordered } = await getOrderedSlides(zip);
    let ctXml = await zip.file("[Content_Types].xml").async("string");
    ctXml = ensureContentTypeDefault(ctXml, "png", "image/png");
    ctXml = ensureContentTypeDefault(ctXml, "jpg", "image/jpeg");

    // 1. Clean trailing empty ANTES/DESPUÉS template slides
    let removedEmpty = 0;
    if (cleanTrailingEmpty) {
      while (ordered.length > 0) {
        const last = ordered[ordered.length - 1];
        const f = zip.file(last.slidePath);
        if (!f) break;
        const xml = await f.async("string");
        const pics = extractPictures(xml);
        const texts = new Set(extractTexts(xml));
        if (pics.length === 0 && texts.size === 2 && texts.has("ANTES") && texts.has("DESPUÉS")) {
          presXml = presXml.replace(last.rawTag, "");
          const relRe = new RegExp(`<Relationship\\b[^>]*Id="${last.rId}"[^>]*\\/?>`);
          presRelsXml = presRelsXml.replace(relRe, "");
          const partName = `/${last.slidePath}`;
          const ovRe = new RegExp(`<Override\\b[^>]*PartName="${partName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*\\/?>`);
          ctXml = ctXml.replace(ovRe, "");
          zip.remove(last.slidePath);
          zip.remove(last.relsPath);
          ordered.pop();
          removedEmpty++;
        } else {
          break;
        }
      }
    }

    let { maxSlideNum, maxMediaNum, maxSldId, maxRId } = findNextNumbers(zip, presXml, presRelsXml);

    const addedSldIdTags = [];
    const addedPresRels = [];
    const addedOverrides = [];

    function registerNewSlide(slideXml, mediaName, mediaBytes) {
      maxSlideNum++;
      maxSldId++;
      maxRId++;
      const slideFilename = `slide${maxSlideNum}.xml`;
      const slidePath = `ppt/slides/${slideFilename}`;
      const relsPath = `ppt/slides/_rels/${slideFilename}.rels`;
      const rId = `rId${maxRId}`;

      zip.file(`ppt/media/${mediaName}`, mediaBytes);
      zip.file(slidePath, slideXml);
      zip.file(relsPath, buildSlideRelsXml(mediaName));

      addedSldIdTags.push(`<p:sldId id="${maxSldId}" r:id="${rId}"/>`);
      addedPresRels.push(`<Relationship Id="${rId}" Type="${REL_SLIDE}" Target="slides/${slideFilename}"/>`);
      addedOverrides.push(`<Override PartName="/${slidePath}" ContentType="${CT_SLIDE}"/>`);
    }

    // 2. Add cover slide
    maxMediaNum++;
    const coverMedia = `image${maxMediaNum}.png`;
    const [cx, cy, cw, ch] = fitBox(cover.width, cover.height, COVER_BOX);
    registerNewSlide(buildCoverSlideXml("rId2", cx, cy, cw, ch), coverMedia, cover.bytes);

    // 3. Add ANTES photo slides
    for (const p of photos) {
      maxMediaNum++;
      const photoMedia = `image${maxMediaNum}.jpg`;
      const [px, py, pw, ph] = fitBox(p.width, p.height, ANTES_BOX);
      registerNewSlide(buildAntesSlideXml("rId2", px, py, pw, ph), photoMedia, p.bytes);
    }

    presXml = presXml.replace("</p:sldIdLst>", addedSldIdTags.join("") + "</p:sldIdLst>");
    presRelsXml = presRelsXml.replace("</Relationships>", addedPresRels.join("") + "</Relationships>");
    ctXml = ctXml.replace("</Types>", addedOverrides.join("") + "</Types>");

    zip.file("ppt/presentation.xml", presXml);
    zip.file("ppt/_rels/presentation.xml.rels", presRelsXml);
    zip.file("[Content_Types].xml", ctXml);

    return {
      added_photo_slides: photos.length,
      removed_empty_slides: removedEmpty,
      total_slides: ordered.length + 1 + photos.length,
    };
  }

  /**
   * Update DESPUÉS photos and/or notes on existing slides.
   * updates: Array<{ slide_path, rels_path, photo?: { bytes, width, height }, note?: string }>
   */
  async function updateDespuesSlides(zip, updates) {
    let ctXml = await zip.file("[Content_Types].xml").async("string");
    ctXml = ensureContentTypeDefault(ctXml, "jpg", "image/jpeg");

    let maxMediaNum = 0;
    zip.forEach((relPath) => {
      const mm = /^ppt\/media\/image(\d+)\./.exec(relPath);
      if (mm) maxMediaNum = Math.max(maxMediaNum, Number(mm[1]));
    });

    let updatedCount = 0;
    const halfW = SLIDE_W / 2;

    for (const item of updates) {
      const hasPhoto = Boolean(item.photo);
      const note = (item.note || "").trim();
      if (!hasPhoto && !note) continue;

      const slideFile = zip.file(item.slide_path);
      if (!slideFile) continue;
      let slideXml = await slideFile.async("string");
      const relsFile = zip.file(item.rels_path);
      let relsXml = relsFile
        ? await relsFile.async("string")
        : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

      // Find max shape id in this slide
      let maxShapeId = 1000;
      const spIdRe = /<p:cNvPr\b[^>]*\bid="(\d+)"/g;
      let m;
      while ((m = spIdRe.exec(slideXml)) !== null) {
        maxShapeId = Math.max(maxShapeId, Number(m[1]));
      }

      const additions = [];

      if (hasPhoto) {
        // Remove any existing right-side pictures
        for (const pic of extractPictures(slideXml)) {
          if (pic.x + pic.cx / 2 >= halfW) {
            slideXml = slideXml.replace(pic.raw, "");
          }
        }

        // Next rId in slide rels
        let maxSlideRId = 0;
        const rRe = /\bId="rId(\d+)"/g;
        while ((m = rRe.exec(relsXml)) !== null) {
          maxSlideRId = Math.max(maxSlideRId, Number(m[1]));
        }
        const newRId = `rId${maxSlideRId + 1}`;

        maxMediaNum++;
        const mediaName = `image${maxMediaNum}.jpg`;
        zip.file(`ppt/media/${mediaName}`, item.photo.bytes);
        relsXml = relsXml.replace(
          "</Relationships>",
          `<Relationship Id="${newRId}" Type="${REL_IMAGE}" Target="../media/${mediaName}"/></Relationships>`
        );

        const box = note
          ? [DESPUES_BOX[0], DESPUES_BOX[1], DESPUES_BOX[2], DESPUES_BOX[3] - 620000]
          : DESPUES_BOX;
        const [px, py, pw, ph] = fitBox(item.photo.width, item.photo.height, box);
        maxShapeId++;
        additions.push(buildPicXml(maxShapeId, "Foto DESPUÉS", newRId, px, py, pw, ph));
      }

      if (note) {
        // Remove previous custom note shapes (any p:sp with text other than ANTES/DESPUÉS)
        const spRe = /<p:sp\b[\s\S]*?<\/p:sp>/g;
        const toRemove = [];
        while ((m = spRe.exec(slideXml)) !== null) {
          const ts = extractTexts(m[0]);
          if (ts.length > 0 && ts.some((t) => t !== "ANTES" && t !== "DESPUÉS")) {
            toRemove.push(m[0]);
          }
        }
        for (const oldSp of toRemove) {
          slideXml = slideXml.replace(oldSp, "");
        }
        maxShapeId++;
        additions.push(buildNoteXml(maxShapeId, note, hasPhoto));
      }

      slideXml = slideXml.replace("</p:spTree>", additions.join("") + "</p:spTree>");
      zip.file(item.slide_path, slideXml);
      zip.file(item.rels_path, relsXml);
      updatedCount++;
    }

    zip.file("[Content_Types].xml", ctXml);
    return { updated_slides: updatedCount };
  }

  return {
    inspectZip,
    appendWalk,
    updateDespuesSlides,
    fitBox,
  };
});
