// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const WIZARD_COMPLETION_ERROR =
    "Failed to save your configuration. Review the previous steps and try again.";

export function adjacentWizardStep(
    currentIndex: number,
    direction: 1 | -1,
    totalSteps: number
): number | undefined {
    const index = currentIndex + direction;
    return index >= 0 && index < totalSteps ? index : undefined;
}

export function getPersistableWizardContext(
    wizardContext: Record<string, any>
): Record<string, any> {
    return Object.fromEntries(
        Object.entries(wizardContext).filter(
            ([key]) => {
                const normalizedKey = key.toLowerCase();
                return !normalizedKey.endsWith("connectionstring")
                    && !normalizedKey.endsWith("creating")
                    && normalizedKey !== "hubassignmentintent"
                    && normalizedKey !== "hubassignmenthubsummary"
                    && normalizedKey !== "hubsourceanalysis";
            }
        )
    );
}

export async function completeWizardSafely(
    onComplete: ((context: Record<string, any>) => void | Promise<void>) | undefined,
    wizardContext: Record<string, any>
): Promise<string | undefined> {
    try {
        await onComplete?.(wizardContext);
        return undefined;
    } catch (error) {
        console.error("Wizard completion failed:", error);
        return WIZARD_COMPLETION_ERROR;
    }
}
