#!/usr/bin/env node
import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const key = argv[i];
    if (!key.startsWith("--")) continue;
    args[key.slice(2)] = argv[i + 1];
    i += 1;
  }
  return args;
}

async function readImageBlob(filePath) {
  const bytes = await fs.readFile(filePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function detectCount(imageDir) {
  const files = await fs.readdir(imageDir);
  return files.filter((name) => /^slide-\d+\.png$/i.test(name)).length;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const imageDir = args["image-dir"];
  const output = args.output;
  const count = args.count ? Number(args.count) : await detectCount(imageDir);
  const previewDir = args["preview-dir"];
  const workspace = args.workspace || process.cwd();

  if (!imageDir || !output || !Number.isInteger(count) || count < 1) {
    throw new Error("Usage: images_to_pptx.mjs --image-dir <dir> --output <file.pptx> [--count 28] [--preview-dir <dir>] [--workspace <artifact-tool-workspace>]");
  }

  const requireFromWorkspace = createRequire(path.join(workspace, "package.json"));
  const { Presentation, PresentationFile } = requireFromWorkspace("@oai/artifact-tool");

  const presentation = Presentation.create({
    slideSize: { width: 1920, height: 1080 },
  });

  for (let i = 1; i <= count; i += 1) {
    const stem = `slide-${String(i).padStart(2, "0")}`;
    const imagePath = path.join(imageDir, `${stem}.png`);
    await fs.access(imagePath);

    const slide = presentation.slides.add();
    slide.background.fill = "#f3efe4";
    slide.images.add({
      blob: await readImageBlob(imagePath),
      contentType: "image/png",
      alt: `Slide ${i}`,
      fit: "cover",
      position: { left: 0, top: 0, width: 1920, height: 1080 },
    });
  }

  if (previewDir) {
    await fs.mkdir(previewDir, { recursive: true });
    for (const [index, slide] of presentation.slides.items.entries()) {
      const stem = `slide-${String(index + 1).padStart(2, "0")}`;
      await writeBlob(
        path.join(previewDir, `${stem}.png`),
        await presentation.export({ slide, format: "png", scale: 1 }),
      );
    }
  }

  await fs.mkdir(path.dirname(output), { recursive: true });
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(output);
  console.log(output);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
