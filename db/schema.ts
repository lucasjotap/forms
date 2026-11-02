// Intentionally empty by default.
// Add Drizzle tables here when the site actually needs a database.
// See examples/d1/db/schema.ts for an opt-in example.
import { sqliteTable, text } from 'drizzle-orm/sqlite-core';
export const responses = sqliteTable('responses', {
  id: text('id').primaryKey(),
  receivedAt: text('received_at').notNull(),
  business: text('business').notNull(),
  activity: text('activity').notNull(),
  answers: text('answers').notNull(),
});
