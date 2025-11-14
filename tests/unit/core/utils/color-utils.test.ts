import { describe, expect, test } from "bun:test";
import {
  lchToHex,
  interpolateLCH,
  calculateGradientFactor,
} from "../../../../src/core/utils/color-utils";
import type { LCHColor } from "../../../../src/core/types";

describe("color-utils", () => {
  describe("lchToHex", () => {
    test("converts pure white", () => {
      const white: LCHColor = { l: 100, c: 0, h: 0 };
      const hex = lchToHex(white);
      expect(hex).toMatch(/^#[0-9a-f]{6}$/i);
      // White should be close to #ffffff
      expect(hex.toLowerCase()).toBe("#ffffff");
    });

    test("converts pure black", () => {
      const black: LCHColor = { l: 0, c: 0, h: 0 };
      const hex = lchToHex(black);
      expect(hex).toMatch(/^#[0-9a-f]{6}$/i);
      expect(hex.toLowerCase()).toBe("#000000");
    });

    test("converts a mid-tone red", () => {
      const red: LCHColor = { l: 50, c: 100, h: 40 };
      const hex = lchToHex(red);
      expect(hex).toMatch(/^#[0-9a-f]{6}$/i);
      // Should produce a reddish color
    });

    test("converts a blue", () => {
      const blue: LCHColor = { l: 50, c: 80, h: 270 };
      const hex = lchToHex(blue);
      expect(hex).toMatch(/^#[0-9a-f]{6}$/i);
      // Should produce a bluish color
    });

    test("handles edge case: very high chroma", () => {
      const highChroma: LCHColor = { l: 50, c: 200, h: 120 };
      const hex = lchToHex(highChroma);
      expect(hex).toMatch(/^#[0-9a-f]{6}$/i);
      // Should clamp values appropriately
    });
  });

  describe("interpolateLCH", () => {
    test("interpolates at t=0 returns first color", () => {
      const color1: LCHColor = { l: 30, c: 50, h: 0 };
      const color2: LCHColor = { l: 70, c: 100, h: 180 };
      const result = interpolateLCH(color1, color2, 0);
      expect(result).toEqual(color1);
    });

    test("interpolates at t=1 returns second color", () => {
      const color1: LCHColor = { l: 30, c: 50, h: 0 };
      const color2: LCHColor = { l: 70, c: 100, h: 180 };
      const result = interpolateLCH(color1, color2, 1);
      expect(result).toEqual(color2);
    });

    test("interpolates at t=0.5 returns midpoint", () => {
      const color1: LCHColor = { l: 20, c: 40, h: 0 };
      const color2: LCHColor = { l: 60, c: 80, h: 100 };
      const result = interpolateLCH(color1, color2, 0.5);
      expect(result.l).toBe(40);
      expect(result.c).toBe(60);
      expect(result.h).toBe(50);
    });

    test("takes shortest path around hue circle (forward)", () => {
      const color1: LCHColor = { l: 50, c: 50, h: 10 };
      const color2: LCHColor = { l: 50, c: 50, h: 350 };
      const result = interpolateLCH(color1, color2, 0.5);
      // Should go through 0, not through 180
      expect(result.h).toBe(0);
    });

    test("takes shortest path around hue circle (backward)", () => {
      const color1: LCHColor = { l: 50, c: 50, h: 350 };
      const color2: LCHColor = { l: 50, c: 50, h: 10 };
      const result = interpolateLCH(color1, color2, 0.5);
      // Should go through 0, not through 180
      expect(result.h).toBe(0);
    });

    test("clamps t below 0", () => {
      const color1: LCHColor = { l: 30, c: 50, h: 0 };
      const color2: LCHColor = { l: 70, c: 100, h: 180 };
      const result = interpolateLCH(color1, color2, -0.5);
      expect(result).toEqual(color1);
    });

    test("clamps t above 1", () => {
      const color1: LCHColor = { l: 30, c: 50, h: 0 };
      const color2: LCHColor = { l: 70, c: 100, h: 180 };
      const result = interpolateLCH(color1, color2, 1.5);
      expect(result).toEqual(color2);
    });

    test("normalizes hue to [0, 360)", () => {
      const color1: LCHColor = { l: 50, c: 50, h: 340 };
      const color2: LCHColor = { l: 50, c: 50, h: 20 };
      const result = interpolateLCH(color1, color2, 0.5);
      expect(result.h).toBeGreaterThanOrEqual(0);
      expect(result.h).toBeLessThan(360);
    });
  });

  describe("calculateGradientFactor", () => {
    test("returns 0 for value at min", () => {
      const factor = calculateGradientFactor(10, 10, 100);
      expect(factor).toBe(0);
    });

    test("returns 1 for value at max", () => {
      const factor = calculateGradientFactor(100, 10, 100);
      expect(factor).toBe(1);
    });

    test("returns 0.5 for value at midpoint", () => {
      const factor = calculateGradientFactor(55, 10, 100);
      expect(factor).toBe(0.5);
    });

    test("clamps value below min to 0", () => {
      const factor = calculateGradientFactor(5, 10, 100);
      expect(factor).toBe(0);
    });

    test("clamps value above max to 1", () => {
      const factor = calculateGradientFactor(150, 10, 100);
      expect(factor).toBe(1);
    });

    test("handles negative ranges", () => {
      const factor = calculateGradientFactor(0, -50, 50);
      expect(factor).toBe(0.5);
    });

    test("returns 0.5 when min equals max", () => {
      const factor = calculateGradientFactor(42, 42, 42);
      expect(factor).toBe(0.5);
    });

    test("handles decimal values", () => {
      const factor = calculateGradientFactor(2.5, 0, 10);
      expect(factor).toBe(0.25);
    });
  });
});

