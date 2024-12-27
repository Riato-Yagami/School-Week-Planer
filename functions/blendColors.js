function blendColors(color1, color2, ratio) {
    // Convert hex color to RGB
    const hexToRgb = (hex) => {
        const bigint = parseInt(hex.replace("#", ""), 16);
        return {
            r: (bigint >> 16) & 255,
            g: (bigint >> 8) & 255,
            b: bigint & 255,
        };
    };

    // Convert RGB to hex color
    const rgbToHex = (r, g, b) => {
        return ((1 << 24) | (r << 16) | (g << 8) | b).toString(16).slice(1).toUpperCase();
    };

    // Convert both colors to RGB
    const color1Rgb = hexToRgb(color1);
    const color2Rgb = hexToRgb(color2);

    // Blend the colors by ratio
    const r = Math.round(color1Rgb.r * (1 - ratio) + color2Rgb.r * ratio);
    const g = Math.round(color1Rgb.g * (1 - ratio) + color2Rgb.g * ratio);
    const b = Math.round(color1Rgb.b * (1 - ratio) + color2Rgb.b * ratio);

    // Return the blended color in hex format
    return rgbToHex(r, g, b);
}

module.exports = (color1, color2, ratio) => {
    return blendColors(color1, color2, ratio);
};