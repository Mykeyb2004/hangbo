import { FileBlob, PresentationFile } from '@oai/artifact-tool';
import fs from 'node:fs/promises';
const variants = ['1_02', '0_96'];
for (const variant of variants) {
  const pptxPath = `/Users/zhangqijin/PycharmProjects/hangbo/tmp/slides/q1_meeting_probe_${variant}.pptx`;
  const outputPath = `/Users/zhangqijin/PycharmProjects/hangbo/tmp/slides/q1_meeting_probe_${variant}.png`;
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
}
