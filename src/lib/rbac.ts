export type Role = "Admin" | "PBD Planner" | "Reviewer" | "IT Manager" | "Read-Only";

export interface Permission {
    action: string;
    allowedRoles: Role[];
}

export const PERMISSIONS: Permission[] = [
    {
        action: "UPLOAD_BUY_FILE",
        allowedRoles: ["Admin", "PBD Planner", "IT Manager"],
    },
    {
        action: "RUN_WORKFLOW",
        allowedRoles: ["Admin", "PBD Planner", "IT Manager"],
    },
    {
        action: "APPROVE_OUTPUT",
        allowedRoles: ["Admin", "Reviewer", "IT Manager"],
    },
    {
        action: "EDIT_MAPPING_TABLES",
        allowedRoles: ["Admin", "IT Manager"],
    },
    {
        action: "VIEW_AUDIT_LOGS",
        allowedRoles: ["Admin", "IT Manager"],
    },
    {
        action: "TOGGLE_API_MODE",
        allowedRoles: ["IT Manager"],
    },
    {
        action: "DOWNLOAD_OUTPUT",
        allowedRoles: ["Admin", "PBD Planner", "Reviewer", "IT Manager", "Read-Only"],
    },
];

export function canPerform(role: Role, action: string): boolean {
    const permission = PERMISSIONS.find((p) => p.action === action);
    if (!permission) return false;
    return permission.allowedRoles.includes(role);
}
