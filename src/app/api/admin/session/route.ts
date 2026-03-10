import { NextRequest, NextResponse } from "next/server";

const ADMIN_COOKIE = "admin_session";
const ADMIN_PASSWORD = process.env.ADMIN_PANEL_PASSWORD || "admin123";

export async function GET(req: NextRequest) {
    const authenticated = req.cookies.get(ADMIN_COOKIE)?.value === "1";
    return NextResponse.json({ authenticated });
}

export async function POST(req: NextRequest) {
    const body = await req.json();
    const password = String(body?.password || "");

    if (password !== ADMIN_PASSWORD) {
        return NextResponse.json({ error: "Invalid password" }, { status: 401 });
    }

    const response = NextResponse.json({ authenticated: true });
    response.cookies.set({
        name: ADMIN_COOKIE,
        value: "1",
        httpOnly: true,
        sameSite: "lax",
        secure: process.env.NODE_ENV === "production",
        path: "/",
        maxAge: 60 * 60 * 8,
    });

    return response;
}

export async function DELETE() {
    const response = NextResponse.json({ authenticated: false });
    response.cookies.set({
        name: ADMIN_COOKIE,
        value: "",
        httpOnly: true,
        sameSite: "lax",
        secure: process.env.NODE_ENV === "production",
        path: "/",
        maxAge: 0,
    });
    return response;
}
