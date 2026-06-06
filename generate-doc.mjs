import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  Table, TableRow, TableCell, WidthType, BorderStyle,
  AlignmentType, ShadingType
} from "docx";
import { writeFileSync } from "fs";

const doc = new Document({
  sections: [{
    properties: {},
    children: [
      // Title
      new Paragraph({
        text: "Building Digital Twin — Dev Log & Error Fixes",
        heading: HeadingLevel.TITLE,
        spacing: { after: 300 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "Project: AI-powered building digital twin | Next.js 16 · LangChain · OpenAI · Neon PostgreSQL · AWS Amplify", italics: true, color: "555555" })],
        spacing: { after: 400 },
      }),

      // Section 1
      new Paragraph({ text: "1. Local Dev Setup", heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 } }),
      new Paragraph({ text: "Error: npm error Missing script: \"dev\"", heading: HeadingLevel.HEADING_2, spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: "Cause: ", bold: true }), new TextRun("Running npm run dev from C:\\3Vibe1\\digitaltwin instead of the actual Next.js project subfolder building-digital-twin.")] }),
      new Paragraph({ children: [new TextRun({ text: "Fix: ", bold: true }), new TextRun("Added a package.json in the parent directory to delegate to the correct subfolder. Now npm run dev works from the parent folder.")] }),
      new Paragraph({ text: 'package.json scripts: { "dev": "npm run dev --prefix building-digital-twin" }', style: "Code", spacing: { before: 100, after: 200 } }),

      // Section 2
      new Paragraph({ text: "2. Default Next.js Boilerplate Page", heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 } }),
      new Paragraph({ children: [new TextRun({ text: "Problem: ", bold: true }), new TextRun("localhost:3000 showed the default Next.js starter page instead of the app.")] }),
      new Paragraph({ children: [new TextRun({ text: "Cause: ", bold: true }), new TextRun("app/page.tsx still contained the boilerplate template — had never been customised.")] }),
      new Paragraph({ children: [new TextRun({ text: "Fix: ", bold: true }), new TextRun("Replaced page.tsx with the full Building Digital Twin dashboard: left panel (Building Status, floor-by-floor asset cards, health score), right panel (AI chat with text input and suggestions), and clickable asset cards that open live SVG illustrations.")] }),

      // Section 3
      new Paragraph({ text: "3. SSL Certificate Error (Local Windows Dev)", heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 } }),
      new Paragraph({ text: "Error: unable to verify the first certificate | code: UNABLE_TO_VERIFY_LEAF_SIGNATURE", style: "Code", spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: "Cause: ", bold: true }), new TextRun("Node.js on Windows could not verify the OpenAI API SSL certificate. Common on Windows where the system CA store is not used by Node by default.")] }),
      new Paragraph({ children: [new TextRun({ text: "Fix: ", bold: true }), new TextRun("Added at the top of API route files:")] }),
      new Paragraph({ text: 'process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";', style: "Code", spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: "Note: ", bold: true, italics: true }), new TextRun({ text: "For local dev only. Not needed on AWS but harmless.", italics: true })] }),

      // Section 4
      new Paragraph({ text: "4. Building Panel Stuck on \"Loading assets...\"", heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 } }),
      new Paragraph({ children: [new TextRun({ text: "Cause: ", bold: true }), new TextRun("fetchAssets() was calling /api/chat with 'get building status' and trying to JSON.parse the AI text reply. The AI returns natural language, not raw JSON, so the parse always failed silently.")] }),
      new Paragraph({ children: [new TextRun({ text: "Fix: ", bold: true }), new TextRun("Created a dedicated app/api/assets/route.ts endpoint that queries the database directly and returns raw JSON. Updated fetchAssets() to call /api/assets (GET) instead of /api/chat.")] }),

      // Section 5
      new Paragraph({ text: "5. Asset SVG Visual Illustrations", heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 } }),
      new Paragraph({ text: "Added SVG-based live illustrations for assets that update in real time based on database status:", spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: "• Lift/Elevator:", bold: true }), new TextRun(" Elevator shaft with animated doors, status panel, floor lights. Detected by asset name containing 'lift' or 'elevator'.")] }),
      new Paragraph({ children: [new TextRun({ text: "• HVAC:", bold: true }), new TextRun(" Spinning fan with vent slats. Detected by name containing 'hvac', 'air', 'fan', or 'cool'.")] }),
      new Paragraph({ children: [new TextRun({ text: "• Others:", bold: true }), new TextRun(" Generic status icon with colour.")] }),
      new Paragraph({ children: [new TextRun({ text: "Status visuals: ", bold: true }), new TextRun("Operational = Green (closed doors, spinning), Faulty = Red (open doors, pulsing alert ring), Maintenance = Amber (half-open doors, wrench).")] }),

      // Section 6
      new Paragraph({ text: "6. GitHub Push", heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 } }),
      new Paragraph({ text: "git init && git add . && git commit -m 'initial commit'", style: "Code" }),
      new Paragraph({ text: "git remote add origin https://github.com/sanpedrokir/building-digital-twin.git", style: "Code" }),
      new Paragraph({ text: "git branch -M main && git push -u origin main", style: "Code", spacing: { after: 150 } }),
      new Paragraph({ children: [new TextRun({ text: "Note: ", bold: true }), new TextRun(".gitignore already excluded node_modules, .next, and .env* — API keys were never committed to GitHub.")] }),

      // Section 7
      new Paragraph({ text: "7. AWS Amplify Setup", heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 } }),
      new Paragraph({ text: "Steps: AWS Console → Amplify → Create new app → GitHub → select repo → branch main → add environment variables → Deploy.", spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: "Environment variables required: ", bold: true }), new TextRun("OPENAI_API_KEY, DATABASE_URL, LANGSMITH_API_KEY, LANGSMITH_TRACING=true, LANGSMITH_PROJECT, LANGSMITH_ENDPOINT, NODE_TLS_REJECT_UNAUTHORIZED=0")] }),

      // Section 8
      new Paragraph({ text: "8. Amplify Build Error — TypeScript Union Type", heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 } }),
      new Paragraph({ text: "Error: Type error: This expression is not callable. (app/api/chat/route.ts:127)", style: "Code", spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: "Cause: ", bold: true }), new TextRun("TypeScript strict mode in production build rejected calling .invoke() on a union type of three different LangChain tool objects.")] }),
      new Paragraph({ children: [new TextRun({ text: "Fix: ", bold: true }), new TextRun("Cast selectedTool to any before calling .invoke():")] }),
      new Paragraph({ text: "const toolResult = await (selectedTool as any).invoke(call.args);", style: "Code", spacing: { after: 200 } }),

      // Section 9
      new Paragraph({ text: "9. Amplify Runtime Error — DATABASE_URL Not Available", heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 } }),
      new Paragraph({ text: 'Error (from browser Network tab): "No database connection string was provided to `neon()`."', style: "Code", spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: "Cause: ", bold: true }), new TextRun("Amplify stores environment variables in AWS SSM Parameter Store. During build, the log showed: !Failed to set up process.env.secrets — SSM injection failed silently, so DATABASE_URL was not available to the Lambda functions at runtime.")] }),

      new Paragraph({ text: "Fix 1 — Switch to Neon serverless driver:", heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }),
      new Paragraph({ children: [new TextRun("Replaced "), new TextRun({ text: "pg", bold: true }), new TextRun(" (TCP-based, needs WebSocket config in Lambda) with "), new TextRun({ text: "@neondatabase/serverless", bold: true }), new TextRun(" (HTTP-based, works natively in serverless).")] }),
      new Paragraph({ text: "npm install @neondatabase/serverless", style: "Code", spacing: { after: 150 } }),

      new Paragraph({ text: "Fix 2 — Correct db.ts wrapper:", heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: "Important: ", bold: true }), new TextRun("neon() returns rows as a plain array T[], NOT { rows: T[] }. Must wrap it to match the existing pool.query() interface. Also initialise neon() inside the function (not at module level) so DATABASE_URL is read at query time, not at cold-start.")] }),
      new Paragraph({ text: "// app/lib/db.ts", style: "Code" }),
      new Paragraph({ text: "import { neon } from '@neondatabase/serverless';", style: "Code" }),
      new Paragraph({ text: "export const pool = {", style: "Code" }),
      new Paragraph({ text: "  query: async (text, values) => {", style: "Code" }),
      new Paragraph({ text: "    const sql = neon(process.env.DATABASE_URL);", style: "Code" }),
      new Paragraph({ text: "    const rows = await sql.query(text, values);", style: "Code" }),
      new Paragraph({ text: "    return { rows };", style: "Code" }),
      new Paragraph({ text: "  },", style: "Code" }),
      new Paragraph({ text: "};", style: "Code", spacing: { after: 150 } }),

      new Paragraph({ text: "Fix 3 — Write env vars to .env.production at build time:", heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }),
      new Paragraph({ children: [new TextRun("Updated "), new TextRun({ text: "amplify.yml", bold: true }), new TextRun(" preBuild to generate .env.production using Amplify console env vars. This bypasses broken SSM injection and embeds values into the Next.js Lambda bundle:")] }),
      new Paragraph({ text: "preBuild commands:", style: "Code" }),
      new Paragraph({ text: "  cat > .env.production << EOF", style: "Code" }),
      new Paragraph({ text: "  DATABASE_URL=$DATABASE_URL", style: "Code" }),
      new Paragraph({ text: "  OPENAI_API_KEY=$OPENAI_API_KEY", style: "Code" }),
      new Paragraph({ text: "  ... (all other env vars)", style: "Code" }),
      new Paragraph({ text: "  EOF", style: "Code" }),
      new Paragraph({ text: "  npm ci", style: "Code", spacing: { after: 200 } }),

      // Section 10
      new Paragraph({ text: "10. Amplify Build Error — result.rows TypeScript Error", heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 } }),
      new Paragraph({ text: "Error: Property 'rows' does not exist on type 'Record<string, any>[]'. (app/api/assets/route.ts:11)", style: "Code", spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: "Cause: ", bold: true }), new TextRun("sql.query() from @neondatabase/serverless returns T[] (a plain array). The route was accessing result.rows which does not exist on an array.")] }),
      new Paragraph({ children: [new TextRun({ text: "Fix: ", bold: true }), new TextRun("Wrap the result in { rows } inside db.ts (see Fix 2 above). All existing result.rows references in route files work without any further changes.")] }),

      // Section 11 — Database Schema
      new Paragraph({ text: "11. Database Schema (Neon PostgreSQL)", heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 } }),
      new Paragraph({ text: "CREATE TABLE building_assets (", style: "Code" }),
      new Paragraph({ text: "  id           SERIAL PRIMARY KEY,", style: "Code" }),
      new Paragraph({ text: "  asset_name   VARCHAR(255) NOT NULL,", style: "Code" }),
      new Paragraph({ text: "  floor_no     INTEGER NOT NULL,", style: "Code" }),
      new Paragraph({ text: "  status       VARCHAR(50) DEFAULT 'operational',", style: "Code" }),
      new Paragraph({ text: "  last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP,", style: "Code" }),
      new Paragraph({ text: "  temperature  NUMERIC,", style: "Code" }),
      new Paragraph({ text: "  energy_usage NUMERIC", style: "Code" }),
      new Paragraph({ text: ");", style: "Code", spacing: { after: 150 } }),
      new Paragraph({ children: [new TextRun({ text: "Valid status values: ", bold: true }), new TextRun("'operational', 'faulty', 'maintenance'")] }),
      new Paragraph({ children: [new TextRun({ text: "Database provider: ", bold: true }), new TextRun("Neon (serverless PostgreSQL) — connection via DATABASE_URL env var using @neondatabase/serverless HTTP driver.")] }),

      // Final status table
      new Paragraph({ text: "12. Final Working State", heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 200 } }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Feature", bold: true })] })] }),
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Status", bold: true })] })] }),
            ],
          }),
          ...([
            ["Local dev (npm run dev)", "✅ Working from parent directory"],
            ["Building panel with assets", "✅ Loads from /api/assets directly"],
            ["AI chat assistant", "✅ GPT-4o-mini with LangChain tools"],
            ["Asset SVG visualisations", "✅ Click any asset for live visual"],
            ["GitHub repo", "✅ github.com/sanpedrokir/building-digital-twin"],
            ["AWS Amplify deployment", "✅ main.d1yecyqv2nozog.amplifyapp.com"],
            ["Database (Neon PostgreSQL)", "✅ Connected via @neondatabase/serverless"],
            ["Auto-deploy on git push", "✅ Amplify watches main branch"],
          ].map(([feature, status]) =>
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph(feature)] }),
                new TableCell({ children: [new Paragraph(status)] }),
              ],
            })
          )),
        ],
      }),

      // Key lessons
      new Paragraph({ text: "Key Lessons Learned", heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 150 } }),
      new Paragraph({ children: [new TextRun({ text: "1. ", bold: true }), new TextRun("Run npm run dev from the correct subfolder — or add a delegating package.json in the parent.")] }),
      new Paragraph({ children: [new TextRun({ text: "2. ", bold: true }), new TextRun("AI chat returns text, not JSON — always use a dedicated API endpoint for structured data fetches.")] }),
      new Paragraph({ children: [new TextRun({ text: "3. ", bold: true }), new TextRun("Use @neondatabase/serverless for Lambda/serverless environments; pg requires WebSocket config.")] }),
      new Paragraph({ children: [new TextRun({ text: "4. ", bold: true }), new TextRun("Amplify SSM injection can silently fail — use amplify.yml preBuild to write env vars to .env.production as a reliable fallback.")] }),
      new Paragraph({ children: [new TextRun({ text: "5. ", bold: true }), new TextRun("neon() returns T[] not { rows: T[] } — wrap with { rows } to match the standard pool.query() interface.")] }),
      new Paragraph({ children: [new TextRun({ text: "6. ", bold: true }), new TextRun("TypeScript is strict in production builds — union type .invoke() calls must be cast to any.")] }),
    ],
  }],
  styles: {
    paragraphStyles: [
      {
        id: "Code",
        name: "Code",
        basedOn: "Normal",
        run: { font: "Courier New", size: 18, color: "1a1a1a" },
        paragraph: {
          shading: { type: ShadingType.SOLID, color: "f0f0f0" },
          spacing: { before: 60, after: 60 },
          indent: { left: 300 },
        },
      },
    ],
  },
});

const buffer = await Packer.toBuffer(doc);
writeFileSync("DEVLOG.docx", buffer);
console.log("✅ DEVLOG.docx created successfully");
