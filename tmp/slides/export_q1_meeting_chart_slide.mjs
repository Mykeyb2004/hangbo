import { FileBlob, PresentationFile } from "@oai/artifact-tool";
import fs from "node:fs/promises";
import path from "node:path";

const pptxPath = "/Users/zhangqijin/PycharmProjects/hangbo/data/ppt/2026/Q1/Q1满意度报告.pptx";
const outputPath = "/Users/zhangqijin/PycharmProjects/hangbo/tmp/slides/q1_meeting_chart_slide.png";
const slideIndex = 10;

const pptx = await FileBlob.load(pptxPath);
const presentation = await PresentationFile.importPptx(pptx);
const slide = presentation.slides.getItem(slideIndex);

if (!slide) {
  throw new Error(`Slide not found at index ${slideIndex}`);
}

const rendered = await presentation.export({
  slide,
  format: "png",
  scale: 1,
});

await fs.mkdir(path.dirname(outputPath), { recursive: true });
console.log("type", typeof rendered, rendered?.constructor?.name);
console.log("keys", Object.keys(rendered ?? {}));
if (rendered?.save) {
  await rendered.save(outputPath);
} else if (typeof rendered?.arrayBuffer === "function") {
  const buffer = Buffer.from(await rendered.arrayBuffer());
  await fs.writeFile(outputPath, buffer);
} else if (rendered instanceof Uint8Array) {
  await fs.writeFile(outputPath, rendered);
} else if (rendered?.data instanceof Uint8Array) {
  await fs.writeFile(outputPath, rendered.data);
} else if (typeof rendered === "string") {
  await fs.writeFile(outputPath, rendered, "utf-8");
}
console.log(outputPath);
