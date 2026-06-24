import { NextResponse } from "next/server";


export function middleware(request) {
  const token = request.cookies.get("auth_token")?.value;
  const isLoginPage = request.nextUrl.pathname === "/login";
  const isApiLogin = request.nextUrl.pathname === "/api/login";
  const isApiLogout = request.nextUrl.pathname === "/api/logout";

  if (isApiLogin || isApiLogout) return NextResponse.next();

  if (!token && !isLoginPage) {
    return NextResponse.redirect(new URL("/login", request.url));
  }

  if (token && isLoginPage) {
    return NextResponse.redirect(new URL("/", request.url));
  }

  return NextResponse.next();
}

export const config = {
  matcher: ["/((?!_next/static|_next/image|favicon.ico).*)"],
};
