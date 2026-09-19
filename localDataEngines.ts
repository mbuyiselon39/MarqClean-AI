import { parsePhoneNumberFromString, type CountryCode } from "libphonenumber-js/max";
import { z } from "zod";

export type PhoneAudit = {
  input: string;
  valid: boolean;
  country?: string;
  international?: string;
  national?: string;
  type?: string;
  reason?: string;
};

export function auditPhone(value: string, defaultCountry: CountryCode = "ZA"): PhoneAudit {
  const input = value.trim();
  if (!input) return { input, valid: false, reason: "Empty value" };
  const phone = parsePhoneNumberFromString(input, defaultCountry);
  if (!phone) return { input, valid: false, reason: "Could not parse number" };
  return {
    input,
    valid: phone.isValid(),
    country: phone.country,
    international: phone.formatInternational(),
    national: phone.formatNational(),
    type: phone.getType(),
    reason: phone.isValid() ? undefined : "Number is not valid for its detected country",
  };
}

export const leadRowSchema = z.object({
  email: z.string().trim().email().optional().or(z.literal("")),
  phone: z.string().trim().optional().or(z.literal("")),
  company: z.string().trim().min(2).optional().or(z.literal("")),
  website: z.string().trim().optional().or(z.literal("")),
});

export type LeadRow = z.infer<typeof leadRowSchema>;

export function validateLeadRow(row: Record<string, unknown>) {
  const result = leadRowSchema.safeParse({
    email: String(row.email ?? ""),
    phone: String(row.phone ?? ""),
    company: String(row.company ?? ""),
    website: String(row.website ?? ""),
  });
  if (result.success) return { valid: true, errors: [] as string[] };
  return { valid: false, errors: result.error.issues.map((issue) => issue.message) };
}

export function extractWebSignals(html: string, sourceUrl = "") {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const text = doc.body?.textContent?.replace(/\\s+/g, " ").trim() ?? "";
  const title = doc.querySelector("title")?.textContent?.trim() ?? "";
  const description = doc.querySelector('meta[name="description"]')?.getAttribute("content")?.trim() ?? "";
  const emails = [...new Set((html.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}/gi) ?? []).map((v) => v.toLowerCase()))];
  const phones = [...new Set((text.match(/(?:\\+?\\d[\\d().\\s-]{7,}\\d)/g) ?? []).map((v) => v.trim()))];
  const links = [...doc.querySelectorAll("a[href]")].map((a) => (a as HTMLAnchorElement).href).filter(Boolean);
  const social = links.filter((url) => /linkedin\\.com|facebook\\.com|instagram\\.com|x\\.com|twitter\\.com/i.test(url));
  const jsonLd = [...doc.querySelectorAll('script[type="application/ld+json"]')].map((node) => node.textContent?.trim()).filter(Boolean);
  const organization = jsonLd.map((raw) => { try { return JSON.parse(raw!); } catch { return null; } }).flatMap((item) => Array.isArray(item) ? item : [item]).find((item) => item && /Organization|Corporation|LocalBusiness/i.test(String(item["@type"] ?? "")));
  return {
    sourceUrl,
    title,
    description,
    company: organization?.name ?? "",
    industry: organization?.industry ?? "",
    emails: emails.slice(0, 20),
    phones: phones.slice(0, 20),
    social: [...new Set(social)].slice(0, 20),
    links: [...new Set(links)].slice(0, 40),
    textPreview: text.slice(0, 600),
  };
}
