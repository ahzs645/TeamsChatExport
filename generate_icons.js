const sharp = require('sharp');
const fs = require('fs');

async function generateIcons() {
  try {
    // Read the SVG file
    const svgBuffer = fs.readFileSync('./assets/icon16.svg');

    // Generate 48x48 icon
    await sharp(svgBuffer)
      .resize(48, 48)
      .png()
      .toFile('./assets/icon48.png');

    // Generate 128x128 icon
    await sharp(svgBuffer)
      .resize(128, 128)
      .png()
      .toFile('./assets/icon128.png');

    console.log('Icons generated successfully!');
  } catch (error) {
    console.error('Error generating icons:', error);
  }
}

generateIcons(); 