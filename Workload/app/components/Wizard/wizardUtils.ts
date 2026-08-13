export const WIZARD_COMPLETION_ERROR =
    "Failed to save your configuration. Review the previous steps and try again.";

export function getPersistableWizardContext(
    wizardContext: Record<string, any>
): Record<string, any> {
    return Object.fromEntries(
        Object.entries(wizardContext).filter(
            ([key]) => {
                const normalizedKey = key.toLowerCase();
                return !normalizedKey.endsWith("connectionstring")
                    && !normalizedKey.endsWith("creating");
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
