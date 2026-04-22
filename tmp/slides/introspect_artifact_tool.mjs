import { Presentation, PresentationFile, FileBlob } from "@oai/artifact-tool";

const presentation = Presentation.create({
  slideSize: { width: 1280, height: 720 },
});
const slide = presentation.slides.add();

console.log("presentation methods", Object.getOwnPropertyNames(Object.getPrototypeOf(presentation)).sort());
console.log("slides methods", Object.getOwnPropertyNames(Object.getPrototypeOf(presentation.slides)).sort());
console.log("slide methods", Object.getOwnPropertyNames(Object.getPrototypeOf(slide)).sort());
console.log("has import/export", typeof PresentationFile.importPptx, typeof PresentationFile.exportPptx, typeof FileBlob.load);
