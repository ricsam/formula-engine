/**
 * Color utilities for conditional styling
 * Handles LCH color space conversions and interpolation
 */

import type { LCHColor } from "../types";

/**
 * Convert LCH color to hex string
 * LCH uses the CIELCH color space which provides perceptually uniform colors
 */
export function lchToHex(color: LCHColor): string {
  // First convert LCH to LAB
  const { l, c, h } = color;
  const hRad = (h * Math.PI) / 180;
  const a = c * Math.cos(hRad);
  const b = c * Math.sin(hRad);

  // Then convert LAB to XYZ (using D65 illuminant)
  let y = (l + 16) / 116;
  let x = a / 500 + y;
  let z = y - b / 200;

  // Apply inverse f function
  const delta = 6 / 29;
  const deltaCubed = delta * delta * delta;
  
  x = x > delta ? x * x * x : 3 * delta * delta * (x - 4 / 29);
  y = y > delta ? y * y * y : 3 * delta * delta * (y - 4 / 29);
  z = z > delta ? z * z * z : 3 * delta * delta * (z - 4 / 29);

  // Scale by D65 white point
  x *= 0.95047;
  y *= 1.0;
  z *= 1.08883;

  // Convert XYZ to RGB
  let r = x * 3.2406 + y * -1.5372 + z * -0.4986;
  let g = x * -0.9689 + y * 1.8758 + z * 0.0415;
  let bVal = x * 0.0557 + y * -0.204 + z * 1.057;

  // Apply gamma correction
  const gammaCorrect = (val: number) => {
    return val > 0.0031308
      ? 1.055 * Math.pow(val, 1 / 2.4) - 0.055
      : 12.92 * val;
  };

  r = gammaCorrect(r);
  g = gammaCorrect(g);
  bVal = gammaCorrect(bVal);

  // Clamp to [0, 1] and convert to 0-255
  r = Math.max(0, Math.min(1, r)) * 255;
  g = Math.max(0, Math.min(1, g)) * 255;
  bVal = Math.max(0, Math.min(1, bVal)) * 255;

  // Convert to hex
  const toHex = (n: number) => {
    const hex = Math.round(n).toString(16);
    return hex.length === 1 ? "0" + hex : hex;
  };

  return `#${toHex(r)}${toHex(g)}${toHex(bVal)}`;
}

/**
 * Convert hex color string to LCH color
 * @param hex Hex color string (e.g., "#FF0000" or "FF0000")
 * @returns LCH color
 */
export function hexToLch(hex: string): LCHColor {
  // Remove # if present and normalize
  const normalizedHex = hex.startsWith("#") ? hex.slice(1) : hex;
  
  // Parse RGB values
  const r = parseInt(normalizedHex.slice(0, 2), 16) / 255;
  const g = parseInt(normalizedHex.slice(2, 4), 16) / 255;
  const b = parseInt(normalizedHex.slice(4, 6), 16) / 255;

  // Apply inverse gamma correction
  const inverseGamma = (val: number) => {
    return val > 0.04045
      ? Math.pow((val + 0.055) / 1.055, 2.4)
      : val / 12.92;
  };

  let rLinear = inverseGamma(r);
  let gLinear = inverseGamma(g);
  let bLinear = inverseGamma(b);

  // Convert RGB to XYZ (using sRGB matrix)
  let x = rLinear * 0.4124564 + gLinear * 0.3575761 + bLinear * 0.1804375;
  let y = rLinear * 0.2126729 + gLinear * 0.7151522 + bLinear * 0.0721750;
  let z = rLinear * 0.0193339 + gLinear * 0.1191920 + bLinear * 0.9503041;

  // Normalize by D65 white point
  x /= 0.95047;
  y /= 1.0;
  z /= 1.08883;

  // Convert XYZ to LAB
  const f = (t: number) => {
    const delta = 6 / 29;
    if (t > delta * delta * delta) {
      return Math.cbrt(t);
    }
    return t / (3 * delta * delta) + 4 / 29;
  };

  const fx = f(x);
  const fy = f(y);
  const fz = f(z);

  const l = 116 * fy - 16;
  const a = 500 * (fx - fy);
  const bLab = 200 * (fy - fz);

  // Convert LAB to LCH
  const c = Math.sqrt(a * a + bLab * bLab);
  let h = Math.atan2(bLab, a) * (180 / Math.PI);
  
  // Normalize hue to [0, 360)
  if (h < 0) {
    h += 360;
  }

  return {
    l: Math.max(0, Math.min(100, l)),
    c: Math.max(0, c),
    h: h,
  };
}

/**
 * Interpolate between two LCH colors
 * @param color1 Starting color
 * @param color2 Ending color
 * @param t Interpolation factor (0-1)
 * @returns Interpolated LCH color
 */
export function interpolateLCH(
  color1: LCHColor,
  color2: LCHColor,
  t: number
): LCHColor {
  // Clamp t to [0, 1]
  t = Math.max(0, Math.min(1, t));

  // Interpolate lightness and chroma linearly
  const l = color1.l + (color2.l - color1.l) * t;
  const c = color1.c + (color2.c - color1.c) * t;

  // Interpolate hue taking the shortest path around the circle
  let h1 = color1.h;
  let h2 = color2.h;

  // Find shortest path
  const diff = h2 - h1;
  if (diff > 180) {
    h1 += 360;
  } else if (diff < -180) {
    h2 += 360;
  }

  let h = h1 + (h2 - h1) * t;

  // Normalize hue to [0, 360)
  h = ((h % 360) + 360) % 360;

  return { l, c, h };
}

/**
 * Calculate interpolation factor for a value within a range
 * @param value Current value
 * @param min Minimum value
 * @param max Maximum value
 * @returns Factor between 0 and 1
 */
export function calculateGradientFactor(
  value: number,
  min: number,
  max: number
): number {
  // Handle edge cases
  if (max === min) {
    return 0.5; // If min equals max, return middle value
  }

  // Calculate factor and clamp to [0, 1]
  const factor = (value - min) / (max - min);
  return Math.max(0, Math.min(1, factor));
}

