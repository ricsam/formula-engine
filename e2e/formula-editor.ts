import { expect, test, type Locator, type Page } from "@playwright/test";

const REFERENCE_BACKGROUND = "rgb(255, 247, 237)";
const REFERENCE_OUTLINE = "rgb(249, 115, 22)";

async function focusFormulaEditor(page: Page): Promise<Locator> {
  const editor = page.getByTestId("formula-editor");
  await expect(editor.locator(".monaco-editor")).toBeVisible({ timeout: 15_000 });

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

test.describe("Formula editor spreadsheet integration", () => {
  test.beforeEach(async ({ page }) => {
    await page.goto("/spreadsheet");
    await expect(page.getByTestId("formula-studio")).toBeVisible();
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!D2",
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
      "Forecast!B2",
    );
    await expectReferenceHighlight(referencedCell);

    await expect(selectedCell).toHaveClass(/rsp-cell-selected/);
    await expect(referencedCell).not.toHaveClass(/rsp-cell-selected/);
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!D2",
    );
  });

  test("caret previews every cell in a resolved range", async ({ page }) => {
    await page.getByTestId("spreadsheet-cell-D7").click();
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!D7",
    );
    await expect(
      page.getByTestId("formula-editor").locator(".view-line"),
    ).toContainText("=SUM(D2:D5)", { timeout: 15_000 });

    const editorInput = await focusFormulaEditor(page);
    await editorInput.press("Home");
    for (let offset = 0; offset < 6; offset++) {
      await editorInput.press("ArrowRight");
    }

    await expect(page.getByTestId("active-reference")).toContainText(
      "Forecast!D2:D5",
    );
    for (const reference of ["D2", "D3", "D4", "D5"]) {
      await expectReferenceHighlight(
        page.getByTestId(`spreadsheet-cell-${reference}`),
      );
    }
    await expect(page.getByTestId("spreadsheet-cell-D7")).toHaveClass(
      /rsp-cell-selected/,
    );
  });

  test("cross-sheet reference previews Assumptions B3 without changing the edited cell", async ({
    page,
  }) => {
    await page.getByTestId("spreadsheet-cell-E2").click();
    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!E2",
    );

    const editor = page.getByTestId("formula-editor");
    await expect(editor.locator(".view-line")).toContainText(
      "=D2*(1-Assumptions!B3)",
      { timeout: 15_000 },
    );
    const editorInput = await focusFormulaEditor(page);
    await editorInput.press("End");
    await editorInput.press("ArrowLeft");

    await expect(page.getByTestId("active-reference")).toContainText(
      "Assumptions!B3",
    );
    await expect(page.getByTestId("visible-sheet")).toHaveText("Assumptions");
    await expectReferenceHighlight(page.getByTestId("spreadsheet-cell-B3"));

    await expect(page.getByTestId("selected-cell-address")).toHaveText(
      "Forecast!E2",
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
    await expect(page.getByTestId("spreadsheet-cell-D2")).toContainText("6,960");
    await expect(page.getByTestId("spreadsheet-cell-E2")).toContainText(
      "6,472.8",
    );
  });
});
