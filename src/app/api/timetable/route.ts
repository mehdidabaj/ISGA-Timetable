import { NextRequest, NextResponse } from "next/server";
import { getTimetable } from "@/lib/timetable";

export const dynamic = "force-dynamic";
export const revalidate = 0;

export async function GET(request: NextRequest) {
  const force = request.nextUrl.searchParams.get("force") === "1";
  const payload = await getTimetable(force);
  return NextResponse.json(payload, {
    headers: {
      "Cache-Control": "no-store, max-age=0",
    },
  });
}
