import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const ROOT = "C:/Users/Administrator/Downloads/agentic-ai-fundamentals/agentic-ai-fundamentals";
const PROJECT = path.join(ROOT, "sta3-github-test/projects/지원소재-매칭-2026-09-15-v01/엔도더마-QC-포트폴리오-2026-09-23-v01");
const BUILD = path.join(PROJECT, ".build");
const OUTPUT = path.join(PROJECT, "output");
const SOURCE_IMAGES = path.join(ROOT, ".tmp/endoderma-poster-2026-09-23-v01/source-images");
const SKILL_DIR = "C:/Users/Administrator/.codex/plugins/cache/openai-primary-runtime/presentations/26.909.12148/skills/Presentations";
const RUNTIME_PYTHON = "C:/Users/Administrator/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe";
const FINAL_PPTX = path.join(OUTPUT, "박기나_학부연구생_실험포스터_엔도더마QC_2026-09-23-v02.pptx");

await fs.mkdir(BUILD, { recursive: true });
await fs.mkdir(OUTPUT, { recursive: true });
const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href);

const W = 1587, H = 1123;
const C = { navy: "#17324D", teal: "#0E6B66", mint: "#DDEFEA", pale: "#F4F8F7", warm: "#F4E7D8", ink: "#1F2933", gray: "#5B6773", light: "#D7E1E5", white: "#FFFFFF" };
const FONT = "Malgun Gothic";
const presentation = Presentation.create({ slideSize: { width: W, height: H } });
const slide = presentation.slides.add();
slide.background.fill = C.white;

function rect(left, top, width, height, fill, radius = 0, line = "none") {
  return slide.shapes.add({ geometry: radius ? "roundRect" : "rect", position: { left, top, width, height }, fill,
    line: line === "none" ? { fill: "none", width: 0 } : { fill: line, width: 1 }, ...(radius ? { borderRadius: radius } : {}) });
}
function textBox(text, left, top, width, height, o = {}) {
  const box = slide.shapes.add({ geometry: "textbox", position: { left, top, width, height }, fill: "none", line: { fill: "none", width: 0 } });
  box.text = text;
  box.text.style = { typeface: FONT, fontSize: o.size ?? 20, bold: o.bold ?? false, color: o.color ?? C.ink,
    autoFit: o.autoFit ?? "shrinkText", ...(o.italic ? { italic: true } : {}), ...(o.align ? { alignment: o.align } : {}) };
  return box;
}
async function picture(name, left, top, width, height, fit = "contain", alt = name) {
  const bytes = await fs.readFile(path.join(SOURCE_IMAGES, name));
  return slide.images.add({ blob: bytes, contentType: name.endsWith(".jpeg") ? "image/jpeg" : "image/png", alt, fit, position: { left, top, width, height } });
}
function section(n, title, left, top, width) {
  rect(left, top, 38, 38, C.teal, 19);
  textBox(String(n), left, top + 2, 38, 32, { size: 18, bold: true, color: C.white, align: "center" });
  textBox(title, left + 52, top - 2, width - 52, 44, { size: 25, bold: true, color: C.navy });
  rect(left, top + 48, width, 3, C.teal);
}

rect(0, 0, W, 18, C.teal);
textBox("고분자 합성 조건 최적화와 FT-IR 기반 반응 확인", 70, 48, 1110, 62, { size: 39, bold: true, color: C.navy });
textBox("ATPB 합성 공정 개선 및 폴리우레탄 경화 조건 비교", 72, 116, 1070, 36, { size: 23, color: C.teal });
textBox("박기나  |  홍익대학교 바이오화학공학과\n학부연구생  |  2023.10–2024.12", 1130, 54, 385, 74, { size: 16, bold: true, color: C.navy, align: "right" });
textBox("화장품 품질관리(QC) 지원 포트폴리오", 1130, 132, 385, 28, { size: 15, color: C.gray, align: "right" });
rect(70, 178, 1447, 2, C.light);

const colW = 458, x1 = 70, x2 = 565, x3 = 1060;
rect(545, 220, 2, 742, C.light); rect(1040, 220, 2, 742, C.light);

section(1, "연구 목적과 합성 흐름", x1, 215, colW);
textBox("HTPB 말단기를 단계적으로 변환해 ATPB를 합성하고, 각 단계의 반응 여부를 FT-IR로 확인했습니다. 이후 반응 조건과 정제 방식을 조정해 공정 효율과 결과 재현성을 높이는 방향을 탐색했습니다.", x1, 278, colW, 104, { size: 17 });
rect(x1, 394, colW, 192, C.pale, 14);
await picture("image1.png", x1 + 12, 408, colW - 24, 164, "contain", "HTPB에서 ATPB로 이어지는 3단계 합성 흐름");
textBox("공정 개선 포인트", x1, 606, 220, 32, { size: 21, bold: true, color: C.navy });
rect(x1, 648, colW, 128, C.mint, 14);
textBox("01", x1 + 18, 663, 40, 28, { size: 18, bold: true, color: C.teal });
textBox("HCl 투입량 1 eq → 0.5 eq\n농도: 원액 → 0.5 M", x1 + 62, 658, 176, 72, { size: 16, bold: true, color: C.navy });
rect(x1 + 242, 664, 2, 88, "#B8D7D0");
textBox("02", x1 + 260, 663, 40, 28, { size: 18, bold: true, color: C.teal });
textBox("3단계 반응을 one-pot으로 진행\nFT-IR로 기존 경로와 유사성 확인", x1 + 304, 658, 140, 88, { size: 15, bold: true, color: C.navy });
rect(x1, 795, 132, 132, C.warm, 66); await picture("image8.png", x1 + 8, 803, 116, 116, "cover", "TPPO 침전 관찰 사진");
textBox("불순물 제거", x1 + 151, 796, 290, 30, { size: 20, bold: true, color: C.navy });
textBox("TPPO의 저온 ether 불용성을 이용해 시료를 냉각하고, 석출된 불순물을 분리했습니다.", x1 + 151, 834, 296, 87, { size: 16 });

section(2, "FT-IR로 반응 단계 확인", x2, 215, colW);
textBox("① Tosylation: 약 3300 cm⁻¹의 OH 피크 감소", x2, 276, colW, 28, { size: 17, bold: true, color: C.teal });
rect(x2, 310, colW, 250, "#FAFCFC", 10, C.light); await picture("image3.png", x2 + 8, 318, colW - 16, 234, "contain", "Tosylation 전후 FT-IR 비교");
textBox("② Azidation → Staudinger reduction", x2, 578, colW, 28, { size: 17, bold: true, color: C.teal });
rect(x2, 612, colW, 252, "#FAFCFC", 10, C.light); await picture("image5.png", x2 + 8, 620, colW - 16, 236, "contain", "Azide와 amine 전환을 확인한 FT-IR 비교");
rect(x2, 883, colW, 68, C.pale, 12);
textBox("분석 결과를 단순 보관하지 않고, 다음 반응 조건을 결정하는 근거로 활용했습니다.", x2 + 18, 894, colW - 36, 48, { size: 16, bold: true, color: C.navy });

section(3, "PU 경화 조건 비교", x3, 215, colW);
textBox("용매·촉매·OH:NCO 비율·온도·시간을 바꾸어 경화 여부를 비교했습니다.", x3, 278, colW, 52, { size: 17 });
rect(x3, 342, colW, 207, "#FAFCFC", 10, C.light); await picture("image9.png", x3 + 8, 350, colW - 16, 191, "contain", "폴리우레탄 경화 조건 비교표");
rect(x3, 568, 206, 132, C.pale, 12); await picture("image10.png", x3 + 10, 578, 186, 112, "contain", "경화되지 않은 시료와 경화 시료 비교");
textBox("관찰 결과", x3 + 225, 570, 220, 30, { size: 20, bold: true, color: C.navy });
textBox("NCO 비율을 높이고 촉매를 사용했을 때 경화가 확인되었습니다.", x3 + 225, 608, 220, 82, { size: 16 });
rect(x3, 718, 164, 210, C.pale, 12); await picture("image16.png", x3 + 8, 726, 148, 194, "cover", "제조한 폴리우레탄 시편");
textBox("시행착오에서 얻은 판단", x3 + 184, 718, 260, 32, { size: 20, bold: true, color: C.navy });
textBox("• 말단 이소시아네이트 방식은 경화가 지나치게 빨라 다른 접근이 필요했습니다.\n• Hexanediol 적용 시 더 부드러운 시편을 확인했습니다.\n• 실패 결과도 조건표로 남겨 다음 실험 설계에 반영했습니다.", x3 + 184, 760, 264, 162, { size: 15 });

rect(0, 988, W, 135, C.navy);
textBox("QC 직무로 연결되는 역량", 70, 1012, 320, 36, { size: 24, bold: true, color: C.white });
textBox("원료·공정 조건 관리", 420, 1012, 220, 32, { size: 19, bold: true, color: "#A9E1D5" });
textBox("실험 조건을 표준화하고 변수별 결과를 비교", 420, 1052, 230, 44, { size: 15, color: C.white });
textBox("시험 결과 해석", 700, 1012, 190, 32, { size: 19, bold: true, color: "#A9E1D5" });
textBox("FT-IR 피크 변화를 근거로 반응 단계 확인", 700, 1052, 230, 44, { size: 15, color: C.white });
textBox("이상 원인 추적", 980, 1012, 190, 32, { size: 19, bold: true, color: "#A9E1D5" });
textBox("불순물·미경화 문제의 원인을 좁혀 조건 수정", 980, 1052, 230, 44, { size: 15, color: C.white });
textBox("기록과 협업", 1260, 1012, 190, 32, { size: 19, bold: true, color: "#A9E1D5" });
textBox("실험 기록 정리와 랩미팅을 통한 결과 공유", 1260, 1052, 230, 44, { size: 15, color: C.white });

slide.speakerNotes.textFrame.setText("Source: 학부연구생_실험_정리_박기나.docx. 문서에 기록된 실험만 반영했으며 GPC 직접 사용 경험은 기재하지 않음.");

const candidatePath = path.join(BUILD, "candidate-poster-v02.pptx");
await (await PresentationFile.exportPptx(presentation)).save(candidatePath);
const preview = await presentation.export({ slide, format: "png", scale: 1 });
await fs.writeFile(path.join(BUILD, "poster-preview-v02.png"), new Uint8Array(await preview.arrayBuffer()));
const layout = await slide.export({ format: "layout" });
await fs.writeFile(path.join(BUILD, "poster-layout-v02.json"), await layout.text());

const stagingDir = path.join(BUILD, ".codex-finalizer");
await fs.mkdir(stagingDir, { recursive: true });
const result = await finalizePresentation({
  explicitTotalSlideCount: 1, requiredNativeTableOwnerSlides: [], requiredNativeChartOwnerSlides: [], workspaceDir: PROJECT,
  candidatePath, finalPath: FINAL_PPTX, pythonExecutable: RUNTIME_PYTHON,
  integrityValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", "15116175,10696575", "--validate-heading-fit"],
  fontPolicy: { basis: "design", families: [FONT] }, verifyArtifactToolImport: true,
  receiptPath: path.join(stagingDir, "poster-v02.validation.json"),
});
console.log(JSON.stringify({ finalPptx: FINAL_PPTX, candidatePath, result }, null, 2));
