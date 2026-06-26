import { NextResponse } from "next/server";

export function middleware(request) {
  const token = request.cookies.get("auth_token")?.value;
  const isLoginPage = request.nextUrl.pathname === "/login";
  const isApiLogin = request.nextUrl.pathname === "/api/login";
  const isApiLogout = request.nextUrl.pathname === "/api/logout";
  const isApiData = request.nextUrl.pathname === "/api/data";
  const isApiAnalyse = request.nextUrl.pathname.startsWith("/api/analyse");

  // API routes always pass through
  if (isApiLogin || isApiLogout || isApiData || isApiAnalyse) {
    return NextResponse.next();
  }

  // No token → redirect to login
  if (!token && !isLoginPage) {
    return NextResponse.redirect(new URL("/login", request.url));
  }

  // Has token → redirect away from login
  if (token && isLoginPage) {
    return NextResponse.redirect(new URL("/", request.url));
  }

  return NextResponse.next();
}

export const config = {
  matcher: ["/((?!_next/static|_next/image|favicon.ico).*)"],
};
