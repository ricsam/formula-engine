import { expect, test, type Locator, type Page } from "@playwright/test";

const REFERENCE_BACKGROUND = "rgb(255, 247, 237)";
const REFERENCE_OUTLINE = "rgb(249, 115, 22)";

async function focusFormulaEditor(page: Page): Promise<Locator> {
  const editor = page.getByTestId("formula-editor");
  await expect(editor.locator(".monaco-editor")).toBeVisible({
    timeout: 15_000,
  });

  const input = editor.getByRole("textbox", { name: "Editor content" });
  await expect(input).toBeAttached({ timeout: 15_000 });
  await input.focus();
  return input;
}

async function expectReferenceHighlight(cell: Locator) {
  await expect(cell).toHaveCSS("background-color", REFERENCE_BACKGROUND);
  await expect(cell).toHaveCSS("outline-color", REFERENCE_OUTLINE);
  await expect(cell).toHaveCSS("outline-style", "solid");
}

async function expectFormulaText(page: Page, text: string) {
  await expect(
    page.getByTestId("formula-editor").locator(".view-line")
  ).toContainText(text);
}

async function expectExactFormulaText(page: Page, text: string) {
  await expect(
    page.getByTestId("formula-editor").locator(".view-line")
  ).toHaveText(text);
}

async function replaceFormula(
  page: Page,
  editorInput: Locator,
  formula: string
) {
  await editorInput.press("Home");
  await editorInput.press("Shift+End");
  await page.keyboard.insertText(formula);
}

async function dragBetween(start: Locator, end: Locator, page: Page) {
  const startBox = await start.boundingBox();
  const endBox = await end.boundingBox();
  if (!startBox || !endBox) {
    throw new Error("Expected both drag targets to be visible");
  }

  await page.mouse.move(
    startBox.x + startBox.width / 2,
    startBox.y + startBox.height / 2
  );
  await page.mouse.down();
  await page.mouse.move(
    endBox.x + endBox.width / 2,
    endBox.y + endBox.height / 2,
    { steps: 8 }
  );
}

async function expectReferenceOverlayCovers(
  page: Page,
  references: readonly string[],
  phase: "selecting" | "selected"
) {
  const overlay = page.getByTestId("spreadsheet-reference-selection");
  await expect(overlay).toBeVisible();
  await expect(overlay).toHaveAttribute("data-reference-phase", phase);

  await expect
    .poll(
      async () => {
        const overlayBox = await overlay.boundingBox();
        if (!overlayBox) return false;

        const cellBoxes = await Promise.all(
          references.map((reference) =>
            page.getByTestId(`spreadsheet-cell-${reference}`).boundingBox()
          )
        );
        if (cellBoxes.some((cellBox) => !cellBox)) return false;

        // Allow one pixel for fractional layout coordinates and browser rounding.
        return cellBoxes.every(
          (cellBox) =>
            cellBox &&
            overlayBox.x <= cellBox.x + 1 &&
            overlayBox.y <= cellBox.y + 1 &&
            overlayBox.x + overlayBox.width >= cellBox.x + cellBox.width - 1 &&
            overlayBox.y + overlayBox.height >= cellBox.y + cellBox.height - 1
        );
      },
      { message: `reference overlay should cover ${references.join(", ")}` }
    )
    .toBe(true);
}

test.describe("Formula editor spreadsheet integration", () => {
  test.beforeEach(async ({ page }) => {
    await page.goto("/spreadsheet");
    await expect(page.getByTestId("formula-studio")).toBeVisible();
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!D2"
    );
  });

  test("caret previews B2 without replacing the selected D2 cell", async ({
    page,
  }) => {
    const selectedCell = page.getByTestId("spreadsheet-cell-D2");
    const referencedCell = page.getByTestId("spreadsheet-cell-B2");
    await expect(selectedCell).toHaveClass(/rsp-cell-selected/);

    const editorInput = await focusFormulaEditor(page);
    await editorInput.press("Home");
    await editorInput.press("ArrowRight");
    await editorInput.press("ArrowRight");

    await expect(page.getByTestId("active-reference")).toContainText(
      "Forecast!B2"
    );
    await expectReferenceHighlight(referencedCell);

    await expect(selectedCell).toHaveClass(/rsp-cell-selected/);
    await expect(referencedCell).not.toHaveClass(/rsp-cell-selected/);
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!D2"
    );
  });

  test("caret previews every cell in a resolved range", async ({ page }) => {
    await page.getByTestId("spreadsheet-cell-D7").click();
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!D7"
    );
    await expect(
      page.getByTestId("formula-editor").locator(".view-line")
    ).toContainText("=SUM(D2:D5)", { timeout: 15_000 });

    const editorInput = await focusFormulaEditor(page);
    await editorInput.press("Home");
    for (let offset = 0; offset < 6; offset++) {
      await editorInput.press("ArrowRight");
    }

    await expect(page.getByTestId("active-reference")).toContainText(
      "Forecast!D2:D5"
    );
    for (const reference of ["D2", "D3", "D4", "D5"]) {
      await expectReferenceHighlight(
        page.getByTestId(`spreadsheet-cell-${reference}`)
      );
    }
    await expect(page.getByTestId("spreadsheet-cell-D7")).toHaveClass(
      /rsp-cell-selected/
    );
  });

  test("cross-sheet reference previews Assumptions B3 without changing the edited cell", async ({
    page,
  }) => {
    await page.getByTestId("spreadsheet-cell-E2").click();
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!E2"
    );

    const editor = page.getByTestId("formula-editor");
    await expect(editor.locator(".view-line")).toContainText(
      "=D2*(1-Assumptions!B3)",
      { timeout: 15_000 }
    );
    const editorInput = await focusFormulaEditor(page);
    await editorInput.press("End");
    await editorInput.press("ArrowLeft");

    await expect(page.getByTestId("active-reference")).toContainText(
      "Assumptions!B3"
    );
    await expect(page.getByTestId("visible-sheet")).toHaveText("Assumptions");
    await expectReferenceHighlight(page.getByTestId("spreadsheet-cell-B3"));

    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!E2"
    );
  });

  test("formula edits apply through Monaco and recalculate dependents", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await editorInput.press("Home");
    await editorInput.press("Shift+End");
    await page.keyboard.insertText("=B2*C2*2");

    await expect(page.getByTestId("apply-formula")).toBeEnabled();
    await page.getByTestId("apply-formula").click();

    await expect(page.getByTestId("selected-cell-value")).toHaveText("6960");
    await expect(page.getByTestId("spreadsheet-cell-D2")).toContainText(
      "6,960"
    );
    await expect(page.getByTestId("spreadsheet-cell-E2")).toContainText(
      "6,472.8"
    );
  });

  test("Enter applies in compact mode without inserting a newline", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await replaceFormula(page, editorInput, "=B2*C2*3");
    await editorInput.press("Enter");

    await expect(page.getByTestId("selected-cell-value")).toHaveText("10440");
    await expectExactFormulaText(page, "=B2*C2*3");
    await expect(
      page.getByTestId("formula-editor").locator(".view-line")
    ).toHaveCount(1);
  });

  test("Enter inserts a line in expanded mode and Control+Enter applies", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await page.getByTestId("toggle-formula-editor").click();
    await editorInput.press("End");
    await editorInput.press("Enter");
    await page.keyboard.insertText("+1");

    await expect(
      page.getByTestId("formula-editor").locator(".view-line")
    ).toHaveCount(2);
    await expect(page.getByTestId("selected-cell-value")).toHaveText("3480");

    await editorInput.press("Control+Enter");
    await expect(page.getByTestId("selected-cell-value")).toHaveText("3481");
    await page.getByTestId("toggle-formula-editor").click();
    await expectExactFormulaText(page, "=B2*C2+1");
  });

  test("plain cell text has no formula diagnostics or caret target", async ({
    page,
  }) => {
    await page.getByTestId("spreadsheet-cell-A11").click();
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!A11"
    );
    await expectFormulaText(page, "Try the editor");
    await expect(page.getByTestId("formula-diagnostics")).toHaveText(
      "Text value"
    );
    await expect(page.getByTestId("active-reference")).toContainText(
      "Move the caret onto a reference"
    );
    await expect(
      page
        .getByTestId("formula-editor")
        .locator(".squiggly-error, .squiggly-warning")
    ).toHaveCount(0);
  });

  test("expands valid formulas into pretty mode and compacts on collapse", async ({
    page,
  }) => {
    await page.getByTestId("spreadsheet-cell-D7").click();
    await expectExactFormulaText(page, "=SUM(D2:D5)");

    await page.getByTestId("toggle-formula-editor").click();
    await expect(page.getByTestId("format-formula")).toBeVisible();
    await expect(
      page.getByTestId("formula-editor").locator(".view-line")
    ).toHaveCount(3);
    await expect(page.getByTestId("apply-formula")).toBeDisabled();

    await page.getByTestId("toggle-formula-editor").click();
    await expectExactFormulaText(page, "=SUM(D2:D5)");
    await expect(page.getByTestId("format-formula")).toHaveCount(0);
  });

  test("leaves invalid formulas untouched until manual formatting succeeds", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await replaceFormula(page, editorInput, "=SUM(A1,,B1)");
    await expect(page.getByTestId("formula-diagnostics")).toContainText(
      "error"
    );
    await editorInput.press("Enter");
    await expect(page.getByTestId("selected-cell-value")).toHaveText("3480");
    await expect(page.getByTestId("formula-save-state")).toContainText(
      "Fix formula errors before applying"
    );

    await page.getByTestId("toggle-formula-editor").click();
    await expect(
      page.getByTestId("formula-editor").locator(".view-line")
    ).toHaveCount(1);
    await expectExactFormulaText(page, "=SUM(A1,,B1)");

    await replaceFormula(page, editorInput, "=IF(B2>0,SUM(B2,C2),0)");
    await expect(page.getByTestId("formula-diagnostics")).toHaveText(
      "Syntax valid"
    );
    await page.getByTestId("format-formula").click();
    await expect(
      page.getByTestId("formula-editor").locator(".view-line")
    ).toHaveCount(8);
  });

  test("offers and accepts built-in function autocomplete", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await editorInput.press("Home");
    await editorInput.press("Shift+End");
    await page.keyboard.insertText("=SU");
    await editorInput.press("Control+Space");

    const suggestions = page.locator(".suggest-widget.visible");
    await expect(suggestions).toBeVisible();
    const sumSuggestion = suggestions.locator(".monaco-list-row").filter({
      has: page.locator(".label-name", { hasText: /^SUM$/ }),
    });
    await expect(sumSuggestion).toBeVisible();
    await sumSuggestion.click();

    await expectFormulaText(page, "=SUM(value1)");
  });

  test("clicking a cell replaces the Monaco selection with its reference", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await editorInput.press("Home");
    await editorInput.press("ArrowRight");
    await editorInput.press("Shift+ArrowRight");
    await editorInput.press("Shift+ArrowRight");

    await page.getByTestId("spreadsheet-cell-C3").click();

    await expectFormulaText(page, "=C3*C2");
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!D2"
    );
    await expect(editorInput).toBeFocused();
    await expectReferenceHighlight(page.getByTestId("spreadsheet-cell-C3"));
    await expect(page.getByTestId("spreadsheet-cell-D2")).toHaveClass(
      /rsp-cell-selected/
    );
    await expect(page.getByTestId("spreadsheet-cell-C3")).not.toHaveClass(
      /rsp-cell-selected/
    );
    await expectReferenceOverlayCovers(page, ["C3"], "selected");

    // A second pick replaces the active picked reference, as it does in Excel.
    await page.getByTestId("spreadsheet-cell-B4").click();
    await expectFormulaText(page, "=B4*C2");
    await expectReferenceHighlight(page.getByTestId("spreadsheet-cell-B4"));
    await expect(page.getByTestId("spreadsheet-cell-D2")).toHaveClass(
      /rsp-cell-selected/
    );
    await expectReferenceOverlayCovers(page, ["B4"], "selected");
  });

  test("clicking a cell replaces the reference under an empty caret", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await editorInput.press("Home");
    await editorInput.press("ArrowRight");
    await editorInput.press("ArrowRight");
    await expect(page.getByTestId("active-reference")).toContainText(
      "Forecast!B2"
    );

    await page.getByTestId("spreadsheet-cell-C3").click();

    await expectExactFormulaText(page, "=C3*C2");
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!D2"
    );
    await expectReferenceOverlayCovers(page, ["C3"], "selected");
  });

  test("sheet tabs retain the formula draft and qualify picked references", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await replaceFormula(page, editorInput, "=SUM()");
    await editorInput.press("ArrowLeft");

    await page.getByTestId("sheet-tab-Assumptions").click();

    await expect(page.getByTestId("visible-sheet")).toHaveText("Assumptions");
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!D2"
    );
    await expectExactFormulaText(page, "=SUM()");

    await page.getByTestId("spreadsheet-cell-B3").click();

    await expectExactFormulaText(page, "=SUM(Assumptions!B3)");
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!D2"
    );
    await expect(editorInput).toBeFocused();
    await expectReferenceOverlayCovers(page, ["B3"], "selected");
  });

  test("Escape leaves reference-picking mode so grid clicks select cells normally", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await editorInput.press("Escape");
    await page.getByTestId("spreadsheet-cell-C3").click();

    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!C3"
    );
    await expectFormulaText(page, "99");
  });

  test("cancelling an active reference drag cannot re-enter picking mode", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await replaceFormula(page, editorInput, "=SUM()");
    await editorInput.press("ArrowLeft");

    await page.getByRole("button", { name: "Zoom In" }).click();
    await dragBetween(
      page.getByTestId("spreadsheet-cell-B2"),
      page.getByTestId("spreadsheet-cell-B2"),
      page
    );
    await expectExactFormulaText(page, "=SUM(B2)");
    await expectReferenceOverlayCovers(page, ["B2"], "selecting");

    await page.keyboard.press("Escape");
    await page.mouse.up();
    await expect(
      page.getByTestId("spreadsheet-reference-selection")
    ).toHaveCount(0);

    await page.getByTestId("spreadsheet-cell-C3").click();
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!C3"
    );
    await expectExactFormulaText(page, "99");
  });

  test("caret movement and manual edits clear a committed reference overlay", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await editorInput.press("Home");
    await editorInput.press("ArrowRight");
    await editorInput.press("Shift+ArrowRight");
    await editorInput.press("Shift+ArrowRight");
    await page.getByTestId("spreadsheet-cell-C3").click();
    await expectReferenceOverlayCovers(page, ["C3"], "selected");

    await editorInput.press("End");
    await expect(
      page.getByTestId("spreadsheet-reference-selection")
    ).toHaveCount(0);

    await replaceFormula(page, editorInput, "=SUM()");
    await editorInput.press("ArrowLeft");
    await page.getByTestId("spreadsheet-cell-B4").click();
    await expectReferenceOverlayCovers(page, ["B4"], "selected");

    await page.keyboard.insertText("+1");
    await expectExactFormulaText(page, "=SUM(B4+1)");
    await expect(
      page.getByTestId("spreadsheet-reference-selection")
    ).toHaveCount(0);
  });

  test("dragging a grid range updates one reference insertion live", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await editorInput.press("Home");
    await editorInput.press("Shift+End");
    await page.keyboard.insertText("=SUM()");
    await editorInput.press("ArrowLeft");

    const start = await page.getByTestId("spreadsheet-cell-B2").boundingBox();
    const end = await page.getByTestId("spreadsheet-cell-C4").boundingBox();
    if (!start || !end)
      throw new Error("Expected spreadsheet cells to be visible");

    await page.mouse.move(
      start.x + start.width / 2,
      start.y + start.height / 2
    );
    await page.mouse.down();
    await page.mouse.move(end.x + end.width / 2, end.y + end.height / 2, {
      steps: 8,
    });

    // The editor is updated before mouseup, while the grid selection is still active.
    await expectFormulaText(page, "=SUM(B2:C4)");
    await expectReferenceOverlayCovers(
      page,
      ["B2", "B3", "B4", "C2", "C3", "C4"],
      "selecting"
    );
    await page.mouse.up();

    await expectFormulaText(page, "=SUM(B2:C4)");
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!D2"
    );
    await expect(editorInput).toBeFocused();
    await expect(page.getByTestId("spreadsheet-cell-D2")).toHaveClass(
      /rsp-cell-selected/
    );
    await expect(page.getByTestId("spreadsheet-cell-B2")).not.toHaveClass(
      /rsp-cell-selected/
    );
    await expectReferenceOverlayCovers(
      page,
      ["B2", "B3", "B4", "C2", "C3", "C4"],
      "selected"
    );
    for (const reference of ["B2", "B3", "B4", "C2", "C3", "C4"]) {
      await expectReferenceHighlight(
        page.getByTestId(`spreadsheet-cell-${reference}`)
      );
    }
  });

  test("cell-to-header drags preserve the finite start of open ranges", async ({
    page,
  }) => {
    const editorInput = await focusFormulaEditor(page);
    await replaceFormula(page, editorInput, "=SUM()");
    await editorInput.press("ArrowLeft");

    await dragBetween(
      page.getByTestId("spreadsheet-cell-B2"),
      page.getByTestId("spreadsheet-row-header-4"),
      page
    );
    await expectExactFormulaText(page, "=SUM(B2:4)");
    await expectReferenceOverlayCovers(page, ["B2", "E4"], "selecting");
    await page.mouse.up();
    await expectReferenceOverlayCovers(page, ["B2", "E4"], "selected");

    await replaceFormula(page, editorInput, "=SUM()");
    await editorInput.press("ArrowLeft");
    await dragBetween(
      page.getByTestId("spreadsheet-cell-B2"),
      page.getByTestId("spreadsheet-col-header-D"),
      page
    );
    await expectExactFormulaText(page, "=SUM(B2:D)");
    await expectReferenceOverlayCovers(page, ["B2", "D11"], "selecting");
    await page.mouse.up();
    await expectReferenceOverlayCovers(page, ["B2", "D11"], "selected");
  });

  test("Control+Arrow navigates to sparse data and reveals an off-screen cell", async ({
    page,
  }) => {
    const target = page.getByTestId("spreadsheet-cell-D40");
    await expect(target).toHaveCount(0);

    await page.getByTestId("spreadsheet-cell-D9").click();
    await page.keyboard.press("Control+ArrowDown");

    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!D40"
    );
    await expect(target).toBeVisible();
    await expect(target).toHaveClass(/rsp-cell-selected/);

    const gridBox = await page.getByTestId("formula-workbook").boundingBox();
    const targetBox = await target.boundingBox();
    if (!gridBox || !targetBox) {
      throw new Error("Expected the workbook and navigated cell to be visible");
    }
    expect(targetBox.y).toBeGreaterThanOrEqual(gridBox.y);
    expect(targetBox.y + targetBox.height).toBeLessThanOrEqual(
      gridBox.y + gridBox.height + 1
    );
  });
});
