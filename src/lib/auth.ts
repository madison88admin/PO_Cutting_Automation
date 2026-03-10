import { NextRequest, NextResponse } from "next/server";
import { Role, canPerform } from "@/lib/rbac";
import { logEvent } from "@/lib/audit";
import { getUser } from "@/lib/db/users";

const ADMIN_SESSION_COOKIE = "admin_session";

export async function getSession(req: NextRequest) {
    const hasAdminSession = req.cookies.get(ADMIN_SESSION_COOKIE)?.value === "1";

    // Anonymous session for public workflow endpoints.
    if (!hasAdminSession) {
        return {
            userId: "anonymous",
            role: "Read-Only" as Role,
            isAuthenticated: false,
        };
    }

    // Admin session for protected control-center endpoints.
    const defaultAdminId = "00000000-0000-0000-0000-000000000001"; // Placeholder
    const user = await getUser(defaultAdminId);

    return {
        userId: user?.id || defaultAdminId,
        role: user?.role || "Admin" as Role,
        isAuthenticated: true,
    };
}

export async function withAuth(
    req: NextRequest,
    action: string,
    handler: (req: NextRequest, session: any) => Promise<NextResponse>
) {
    const session = await getSession(req);

    if (!canPerform(session.role, action)) {
        await logEvent({
            eventName: "UNAUTHORIZED_ACCESS",
            userId: session.userId,
            timestamp: new Date().toISOString(),
            details: `Attempted action: ${action}`,
        });

        return NextResponse.json(
            { error: "Forbidden", message: "Unauthorized action" },
            { status: 403 }
        );
    }

    return handler(req, session);
}
