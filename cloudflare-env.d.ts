declare namespace Cloudflare {
  interface Env {
    DB?: D1Database;
    BUCKET?: R2Bucket;
    ADMIN_EMAIL?: string;
    ADMIN_PASSWORD?: string;
    TRUST_PLATFORM_AUTH?: string;
    FORM_SECRET?: string;
  }
}
