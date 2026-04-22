import { FileBlob, PresentationFile } from '@oai/artifact-tool';
import fs from 'node:fs/promises';
const pptxPath = '/Users/zhangqijin/PycharmProjects/hangbo/tmp/slides/q1_meeting_textbox_probe.pptx';
const outputPath = '/Users/zhangqijin/PycharmProjects/hangbo/tmp/slides/q1_meeting_textbox_probe.png';
const pptx = await FileBlob.load(pptxPath);
const presentation = await PresentationFile.importPptx(pptx);
const slide = presentation.slides.getItem(presentation.slides.count - 1);
const rendered = await presentation.export({ slide, format: 'png', scale: 1 });
if (rendered?.save) {
  await rendered.save(outputPath);
} else {
  const buffer = Buffer.from(await rendered.arrayBuffer());
  await fs.writeFile(outputPath, buffer);
}
console.log(outputPath);
