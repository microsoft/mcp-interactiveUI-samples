import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import dotenv from "dotenv";
dotenv.config();
import { ensureTables, getExpensesTable } from "./db.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DB_DIR = path.resolve(__dirname, "..", "db");

interface SeedExpense {
  id: string;
  description: string;
  category: string;
  merchant: string;
  date: string;
  amount: number;
  currency: string;
  card_last_four: string;
  status: string;
}

async function seed() {
  console.log("🌱 Seeding Expense Submission database...\n");
  await ensureTables();

  console.log("📋 Seeding Expenses...");
  const raw = fs.readFileSync(path.join(DB_DIR, "expenses.json"), "utf-8");
  const expenses: SeedExpense[] = JSON.parse(raw);

  for (const e of expenses) {
    await getExpensesTable().upsertEntity(
      {
        partitionKey: "expenses",
        rowKey: e.id,
        description: e.description,
        category: e.category,
        merchant: e.merchant,
        date: e.date,
        amount: e.amount,
        currency: e.currency,
        cardLastFour: e.card_last_four,
        status: e.status,
        businessPurpose: "",
        projectCode: "",
        attendees: "",
        notes: "",
      },
      "Replace"
    );
    console.log(`  ✓ ${e.id} - ${e.merchant}`);
  }

  console.log("\n✅ Seeding complete!");
}

seed().catch(console.error);
