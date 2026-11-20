import { NextRequest, NextResponse } from "next/server";
import { ingestExcelUpload } from "@/lib/timetable";

export const dynamic = "force-dynamic";
export const revalidate = 0;

const MAX_BYTES = 12 * 1024 * 1024;

export async function POST(request: NextRequest) {
  const form = await request.formData();
  const file = form.get("file");

  if (!(file instanceof File)) {
    return NextResponse.json(
      { error: "Ajoutez un fichier Excel (.xlsx)." },
      { status: 400 },
    );
  }

  const name = file.name.toLowerCase();
  if (!name.endsWith(".xlsx") && !name.endsWith(".xls")) {
    return NextResponse.json(
      { error: "Format non supporté. Utilisez un fichier .xlsx." },
      { status: 400 },
    );
  }

  if (file.size <= 0 || file.size > MAX_BYTES) {
    return NextResponse.json(
      { error: "Fichier trop volumineux (max 12 Mo) ou vide." },
      { status: 400 },
    );
  }

  try {
    const buffer = Buffer.from(await file.arrayBuffer());
    // Basic zip/xlsx signature check for .xlsx
    if (name.endsWith(".xlsx") && !(buffer[0] === 0x50 && buffer[1] === 0x4b)) {
      return NextResponse.json(
        { error: "Le fichier ne semble pas être un Excel valide." },
        { status: 400 },
      );
    }
    const payload = await ingestExcelUpload(buffer);
    return NextResponse.json(payload);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Import impossible";
    return NextResponse.json({ error: message }, { status: 422 });
  }
}
