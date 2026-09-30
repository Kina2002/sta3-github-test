import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { pathToFileURL } from "node:url";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const ROOT = "C:/Users/Administrator/Downloads/agentic-ai-fundamentals/agentic-ai-fundamentals";
const PROJECT = path.join(ROOT, "sta3-github-test/projects/지원소재-매칭-2026-09-15-v01/엔도더마-QC-포트폴리오-2026-09-23-v01");
const BUILD = path.join(PROJECT, ".build");
const OUTPUT = path.join(PROJECT, "output");
const SOURCE = "C:/Users/Administrator/Desktop/박기나/제출 서류/박기나_학부연구생_실험포스터 수정 1차.pptx";
const SOURCE_IMAGES = path.join(ROOT, ".tmp/endoderma-poster-2026-09-23-v01/source-images");
const SKILL_DIR = "C:/Users/Administrator/.codex/plugins/cache/openai-primary-runtime/presentations/26.909.12148/skills/Presentations";
const RUNTIME_PYTHON = "C:/Users/Administrator/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe";
const FINAL_PPTX = path.join(OUTPUT, "박기나_학부연구생_실험포스터_수정2차_2026-09-23.pptx");

await fs.mkdir(BUILD, { recursive: true });
await fs.mkdir(OUTPUT, { recursive: true });
const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href);

const sourceBytes = await fs.readFile(SOURCE);
const sourceSha256 = crypto.createHash("sha256").update(sourceBytes).digest("hex");
const deck = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const slide = deck.slides.items[0];

const COLORS = { navy: "#17324D", teal: "#0E6B66", pale: "#F4F8F7", light: "#D7E1E5", white: "#FFFFFF", ink: "#1F2933", gray: "#5B6773" };
const FONT = "Malgun Gothic";

function addRect(left, top, width, height, fill, radius = 0, line = "none") {
  return slide.shapes.add({
    geometry: radius ? "roundRect" : "rect",
    position: { left, top, width, height },
    fill,
    line: line === "none" ? { fill: "none", width: 0 } : { fill: line, width: 1 },
    ...(radius ? { borderRadius: radius } : {}),
  });
}

function addText(text, left, top, width, height, opts = {}) {
  const box = slide.shapes.add({
    geometry: "textbox",
    position: { left, top, width, height },
    fill: "none",
    line: { fill: "none", width: 0 },
  });
  box.text = text;
  box.text.style = {
    typeface: FONT,
    fontSize: opts.size ?? 17,
    bold: opts.bold ?? false,
    color: opts.color ?? COLORS.ink,
    autoFit: "shrinkText",
  };
  return box;
}

async function addImage(name, left, top, width, height, fit = "contain", alt = name, crop) {
  const bytes = await fs.readFile(path.join(SOURCE_IMAGES, name));
  return slide.images.add({
    blob: bytes,
    contentType: name.endsWith(".jpeg") ? "image/jpeg" : "image/png",
    alt,
    fit,
    position: { left, top, width, height },
    ...(crop ? { crop } : {}),
  });
}

// Keep the user's other edits. Replace only the content of section 3.
deck.resolve("sh/e1oz61c3").text.replace("PU 경화 조건 비교", "Hexanediol chain extender 적용");
deck.resolve("sh/036h8but").text.replace(
  "용매·촉매·OH:NCO 비율·온도·시간을 바꾸어 경화 여부를 비교 및 시편 개선",
  "Hexanediol을 첨가해 기존 PDI 기반 PU보다 부드러운 시편을 제조했습니다.",
);

for (const id of [
  "sh/nqxgjqtk", "im/8z6tofq5", "sh/9sfylgbq", "im/ru1crep4",
  "sh/buh0nqtg", "sh/at8zelcb", "sh/6x8f2583", "sh/7yhgbq9o",
]) {
  deck.resolve(id).delete();
}

addText("실험 절차", 1060, 341, 180, 30, { size: 20, bold: true, color: COLORS.navy });
addRect(1060, 378, 458, 99, COLORS.pale, 12);
addText(
  "1  HTPB와 Hexanediol을 진공 건조한 뒤 질소 퍼징\n2  IPDI 또는 PDI를 넣어 반응\n3  촉매를 넣고 짧게 반응한 뒤 틀에 주입\n4  오븐에서 경화",
  1078, 389, 424, 78, { size: 15.5, color: COLORS.ink },
);

addText("조건별 경화 결과", 1060, 493, 220, 30, { size: 20, bold: true, color: COLORS.navy });
addRect(1060, 530, 458, 181, COLORS.white, 10, COLORS.light);
await addImage("image15.png", 1068, 538, 442, 165, "contain", "Hexanediol 첨가 조건과 경화 여부를 정리한 표");

addText("시편 결과", 1060, 728, 160, 30, { size: 20, bold: true, color: COLORS.navy });
addRect(1060, 766, 176, 178, COLORS.pale, 12);
await addImage("image16.png", 1068, 774, 160, 162, "cover", "Hexanediol을 적용해 제조한 폴리우레탄 시편");

addText("KP-01-28", 1258, 766, 240, 28, { size: 17, bold: true, color: COLORS.teal });
addText(
  "Hexanediol을 chain extender로 적용한 시편에서 기존 PDI 기반 PU보다 부드러운 특성을 확인했습니다.",
  1258, 801, 245, 118, { size: 16, color: COLORS.ink },
);

slide.speakerNotes.textFrame.setText(
  "Section 3 source: 학부연구생_실험_정리_박기나.docx paragraphs 114-126 and supplied screenshot. Hexamethylenediamine comparison was planned but not completed, so it is not presented as a result.",
);

const candidatePath = path.join(BUILD, "candidate-user-revision-2.pptx");
await (await PresentationFile.exportPptx(deck)).save(candidatePath);
const preview = await deck.export({ slide, format: "png", scale: 1 });
await fs.writeFile(path.join(BUILD, "user-revision-2-preview.png"), new Uint8Array(await preview.arrayBuffer()));
const layout = await slide.export({ format: "layout" });
await fs.writeFile(path.join(BUILD, "user-revision-2.layout.json"), await layout.text());

const receiptDir = path.join(BUILD, ".codex-finalizer");
await fs.mkdir(receiptDir, { recursive: true });
const result = await finalizePresentation({
  explicitTotalSlideCount: 1,
  requiredNativeTableOwnerSlides: [],
  requiredNativeChartOwnerSlides: [],
  workspaceDir: PROJECT,
  candidatePath,
  finalPath: FINAL_PPTX,
  pythonExecutable: RUNTIME_PYTHON,
  integrityValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", "15116175,10696575", "--validate-heading-fit"],
  fontPolicy: { basis: "reference", families: [FONT, "Calibri", "한컴 말랑말랑 Bold"], referencePath: SOURCE, referenceSha256: sourceSha256 },
  verifyArtifactToolImport: true,
  receiptPath: path.join(receiptDir, "user-revision-2.validation.json"),
});

console.log(JSON.stringify({ finalPptx: FINAL_PPTX, sourceSha256, result }, null, 2));
