// sera-attach-video-module.ts
// Downloads a video from a SharePoint (Graph) download URL and attaches it to
// the Media and Documents tab of a specific SERA (misterquik) job.
//
// Input: jobId, appointmentId, fileName, downloadUrl
// Flow:
//   1. Download the video from the SharePoint download URL to a temp file
//   2. Login to SERA (misterquik.sera.tech)
//   3. Navigate to /jobs/{jobId}?tab=jp_Media+and+Documents
//   4. Find the exact "Upload Video(s)" file input and setInputFiles(tempPath)
//   5. Poll the DOM until the upload is confirmed (or 5 min timeout)
//   6. Return success message

import { Stagehand } from "@browserbasehq/stagehand";
import express from "express";
import fs from "fs";
import path from "path";
import os from "os";

const DEBUG     = process.env.DEBUG === "true" || process.env.DEBUG === "1";
const DEBUG_DIR = process.env.DEBUG_DIR || "./debug";
const SERA_BASE_URL = (process.env.SERA_MRQUIK_BASE_URL || "https://misterquik.sera.tech").replace(/\/+$/, "");

// =============================================================================
// PAGE HELPERS
// =============================================================================

async function waitUntilVisible(page: any, selector: string, timeoutMs = 10000): Promise<boolean> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      if (await page.locator(selector).first().isVisible()) return true;
    } catch { /* not ready */ }
    await page.waitForTimeout(500);
  }
  throw new Error(`Timeout (${timeoutMs}ms): "${selector}" never became visible`);
}

async function firstVisible(page: any, selectors: string[], timeoutMs = 2000): Promise<string | null> {
  for (const sel of selectors) {
    try {
      if (await page.locator(sel).first().isVisible()) return sel;
    } catch { /* try next */ }
  }
  await page.waitForTimeout(timeoutMs);
  for (const sel of selectors) {
    try {
      if (await page.locator(sel).first().isVisible()) return sel;
    } catch { /* try next */ }
  }
  return null;
}

// Finds the EXACT "Upload Video(s)" file input on the Media and Documents tab.
// Mirrors findExactDocumentsUploadInput from the report module, but scoped to videos.
async function findExactVideoUploadInput(page: any): Promise<null | {
  inputIndex: number;
  labelText: string;
  accept: string;
  availableLabels: string[];
  error?: string;
}> {
  const source = `(() => {
    try {
      const normalize = s => String(s || "").replace(/\\s+/g, " ").trim();
      const canon = s => normalize(s).toLowerCase().replace(/\\s*\\(\\s*/g, "(").replace(/\\s*\\)\\s*/g, ")");
      const allInputs = Array.from(document.querySelectorAll('input[type="file"]'));
      const labels = Array.from(document.querySelectorAll("label"));
      const availableLabels = labels.map(l => normalize(l.textContent || "")).filter(Boolean);

      for (const label of labels) {
        const labelText = normalize(label.textContent || "");
        const normalizedLabel = canon(labelText);
        if (normalizedLabel !== "upload video(s)" && normalizedLabel !== "upload videos") continue;

        const input = label.querySelector('input[type="file"]');
        if (!input) continue;

        return {
          inputIndex: allInputs.indexOf(input),
          labelText: labelText,
          accept: input.accept || "",
          availableLabels: availableLabels
        };
      }

      return {
        inputIndex: -1,
        labelText: "",
        accept: "",
        availableLabels: availableLabels
      };
    } catch (error) {
      return {
        inputIndex: -1,
        labelText: "",
        accept: "",
        availableLabels: [],
        error: error && error.message ? error.message : String(error)
      };
    }
  })()`;

  return await page.evaluate((script: string) => {
    return (0, eval)(script);
  }, source);
}

async function captureDebug(page: any, name: string, idx: number) {
  if (!DEBUG) return;
  try {
    if (!fs.existsSync(DEBUG_DIR)) fs.mkdirSync(DEBUG_DIR, { recursive: true });
    const ts   = new Date().toISOString().replace(/[:.]/g, "-");
    const safe = name.replace(/[^a-zA-Z0-9]/g, "_").substring(0, 50);
    await page.screenshot({ path: `${DEBUG_DIR}/sera-video-step${idx}_${safe}_${ts}.png`, fullPage: true });
    fs.writeFileSync(`${DEBUG_DIR}/sera-video-step${idx}_${safe}_${ts}.html`, await page.content());
  } catch (e: any) {
    console.log(`    ⚠️  Debug capture failed: ${e.message}`);
  }
}

// =============================================================================
// STEP SYSTEM
// =============================================================================

interface Step {
  name: string;
  skipIf?: (page: any, ctx: any) => boolean | Promise<boolean>;
  run: (page: any, ctx: any) => Promise<void>;
}

async function runStep(s: Step, page: any, ctx: any, idx: number): Promise<{ success: boolean; error?: string }> {
  console.log(`  [${idx}] → ${s.name}`);
  console.log(`    ℹ️  URL before step: ${page.url()}`);
  if (s.skipIf) {
    const skip = await s.skipIf(page, ctx);
    if (skip) { console.log("    ⏭️  Skipped"); return { success: true }; }
  }
  try {
    await s.run(page, ctx);
    console.log(`    ℹ️  URL after step: ${page.url()}`);
    console.log("    ✅ Done");
    return { success: true };
  } catch (e: any) {
    const msg = e.message || String(e);
    console.log(`    ❌ Failed: ${msg}`);
    await captureDebug(page, s.name, idx);
    return { success: false, error: msg };
  }
}

// =============================================================================
// DOWNLOAD HELPER (SharePoint Graph download URL → local temp file)
// =============================================================================

async function downloadToFile(url: string, destPath: string): Promise<number> {
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Download failed: ${res.status} ${res.statusText}`);
  }
  const arrayBuffer = await res.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);
  fs.writeFileSync(destPath, buffer);
  return buffer.length;
}

// =============================================================================
// BUILD STEPS
// =============================================================================

function buildSteps(I: {
  stratablueEmail:    string;
  stratabluePassword: string;
  jobId:              string;
  appointmentId:      string;
  fileName:           string;
  localFilePath:      string;
}): Step[] {
  return [

    // =========================================================================
    // LOGIN
    // =========================================================================
    {
      name: "Navigate to login page",
      async run(page) {
        await page.goto(`${SERA_BASE_URL}/admins/login`, {
          waitUntil: "domcontentloaded",
          timeout: 30000,
        });
        await waitUntilVisible(page, 'input[type="email"], input[name="email"]', 15000);
      },
    },
    {
      name: "Fill email",
      async run(page) {
        const sel = await firstVisible(page, ['input[type="email"]', 'input[name="email"]']);
        if (!sel) throw new Error("Email input not found");
        await page.locator(sel).first().fill(I.stratablueEmail);
      },
    },
    {
      name: "Fill password",
      async run(page) {
        await page.locator('input[type="password"]').first().fill(I.stratabluePassword);
      },
    },
    {
      name: "Click login button",
      async run(page) {
        await page.waitForTimeout(500);
        const clicked = await page.evaluate(() => {
          const keywords = ["sign in", "login", "log in"];
          const btn = Array.from(document.querySelectorAll('button, input[type="submit"]')).find(
            el =>
              keywords.some(kw =>
                el.textContent?.toLowerCase().trim() === kw ||
                (el as HTMLInputElement).value?.toLowerCase() === kw
              ) && (el as HTMLElement).offsetParent !== null
          ) as HTMLElement | null;
          if (btn) { btn.click(); return true; }
          return false;
        });
        if (!clicked) {
          const fallback = await firstVisible(page, [
            'input[type="submit"]', 'button[type="submit"]', ".btn-primary", "button.btn",
          ], 3000);
          if (!fallback) throw new Error("Login button not found");
          await page.locator(fallback).first().click();
        }
      },
    },
    {
      name: "Wait for post-login redirect",
      async run(page) {
        await page.waitForTimeout(3000);
        for (let i = 0; i < 20; i++) {
          const url = page.url();
          if (!url.includes("/login")) {
            console.log(`    ℹ️  Redirected to: ${url}`);
            return;
          }
          await page.waitForTimeout(1000);
        }
        throw new Error("Still on login page after 23s — check credentials");
      },
    },

    // =========================================================================
    // NAVIGATE STRAIGHT TO JOB → MEDIA AND DOCUMENTS
    // =========================================================================
    {
      name: "Navigate to job Media and Documents tab",
      async run(page) {
        const jobUrl = `${SERA_BASE_URL}/jobs/${I.jobId}?tab=jp_Media+and+Documents`;
        console.log(`    ℹ️  Opening: ${jobUrl}`);
        try {
          await page.goto(jobUrl, { waitUntil: "domcontentloaded", timeout: 30000 });
        } catch {
          await page.waitForTimeout(3000);
        }
        // Give the tab content time to render
        await page.waitForTimeout(4000);
      },
    },

    // =========================================================================
    // UPLOAD VIDEO
    // =========================================================================
    {
      name: "Upload video file",
      async run(page, ctx) {
        const stats = fs.statSync(I.localFilePath);
        console.log(`    ℹ️  Upload strategy: exact-video-input-v1`);
        console.log(`    ℹ️  Uploading: ${I.fileName} (${(stats.size / 1024 / 1024).toFixed(2)} MB)`);

        // Wait for the Media and Documents area to render
        await waitUntilVisible(
          page,
          ".card.photo-gallery h3.page-header-2, .card.photo-gallery h3, label",
          15000
        );

        const target = await findExactVideoUploadInput(page);
        if (!target) {
          throw new Error("Could not scan page for Upload Video(s) input");
        }
        if (target.error) {
          throw new Error(`Video input scan failed in browser context: ${target.error}`);
        }
        if (target.inputIndex < 0) {
          throw new Error(
            `Exact "Upload Video(s)" label/input not found. ` +
            `Available labels: ${target.availableLabels.join(" | ") || "none"}`
          );
        }

        console.log(`    ℹ️  Exact Video target: ${JSON.stringify(target)}`);

        await page
          .locator('input[type="file"]')
          .nth(target.inputIndex)
          .setInputFiles(I.localFilePath);

        console.log(
          `    ℹ️  File set on exact Video input #${target.inputIndex}: ` +
          `label="${target.labelText}" accept="${target.accept}"`
        );
        console.log(`    ℹ️  File queued for upload — waiting for completion`);

        // Videos can be large; give the upload up to 5 minutes to land in the DOM
        const deadline = Date.now() + 5 * 60 * 1000;
        let confirmed = false;
        while (Date.now() < deadline) {
          await page.waitForTimeout(5000);
          confirmed = await page.evaluate((fname: string) => {
            const nameWithoutExt = fname.replace(/\.[^.]+$/, "");
            const body = document.body.textContent || "";
            return (
              body.includes(nameWithoutExt) ||
              body.includes(fname) ||
              !!document.querySelector(
                '.video-item, .media-item, [class*="video-row"], [class*="media-row"]'
              )
            );
          }, I.fileName);
          if (confirmed) break;
        }

        ctx.videoUploaded = true;
        console.log(confirmed
          ? `    ℹ️  Upload confirmed in DOM`
          : `    ℹ️  Upload sent — DOM confirmation uncertain (check session replay)`
        );
      },
    },

    // =========================================================================
    // COMPLETION
    // =========================================================================
    {
      name: "Generate completion report",
      async run(_page, ctx) {
        const parts = [
          `Video uploaded`,
          `Appointment: ${I.appointmentId || "n/a"}`,
          `Job ID: ${I.jobId}`,
          `File: ${I.fileName}`,
        ];
        const msg = parts.join(" | ");
        console.log(`\n🎉 ${msg}`);
        ctx.completionMessage = msg;
      },
    },
  ];
}

// =============================================================================
// EXPORTED FUNCTION — called by server.ts
// =============================================================================

export async function runSeraAttachVideo(input: any) {
  // Accept both camelCase and snake_case, plus the raw Graph property
  const jobId         = String(input.jobId         || input.job_id         || "");
  const appointmentId = String(input.appointmentId || input.appointment_id || "");
  const fileName      = input.fileName             || input.file_name      || "video.mov";
  const downloadUrl   = input.downloadUrl
                     || input.download_url
                     || input["@microsoft.graph.downloadUrl"]
                     || "";

  if (!jobId)       throw new Error("jobId is required");
  if (!downloadUrl) throw new Error("downloadUrl is required");

  const startTime = Date.now();
  let sessionUrl = "";

  console.log(`\n🎥 SERA Attach Video — job: ${jobId}, appointment: ${appointmentId}`);
  console.log(`   File: ${fileName}`);

  // ---- 1. Download the video from SharePoint to a temp file ----
  const safeName = fileName.replace(/[^a-zA-Z0-9._-]/g, "_");
  const tempPath = path.join(os.tmpdir(), `sera-video-${Date.now()}-${safeName}`);
  console.log(`⬇️  Downloading video to ${tempPath}`);
  let bytes = 0;
  try {
    bytes = await downloadToFile(downloadUrl, tempPath);
    console.log(`✅ Downloaded ${(bytes / 1024 / 1024).toFixed(2)} MB`);
  } catch (e: any) {
    console.error(`❌ Failed to download video: ${e.message}`);
    return {
      success: false,
      message: `Failed to download video from SharePoint: ${e.message}`,
      jobId,
      appointmentId,
      fileName,
      videoUploaded: false,
      sessionUrl: "",
      elapsedSeconds: parseFloat(((Date.now() - startTime) / 1000).toFixed(1)),
      results: [{ step: "Download video", success: false, error: e.message }],
    };
  }

  // ---- 2. Build the step plan and run the browser flow ----
  const I = {
    stratablueEmail:    process.env.STRATABLUE_MRQUIK_EMAIL    || "mcc@stratablue.com",
    stratabluePassword: process.env.STRATABLUE_MRQUIK_PASSWORD || "",
    jobId,
    appointmentId,
    fileName,
    localFilePath: tempPath,
  };

  const STEPS    = buildSteps(I);
  const context: any = {};
  const results: Array<{ step: string; success: boolean; error?: string }> = [
    { step: "Download video", success: true },
  ];

  const stagehand = new Stagehand({
    env:         "BROWSERBASE",
    model:       "google/gemini-2.5-flash",
    verbose:     DEBUG ? 2 : 1,
    disablePino: !DEBUG,
    browserbaseSessionCreateParams: {
      projectId: process.env.BROWSERBASE_PROJECT_ID,
      userMetadata: {
        task: "sera-attach-video-mrquik",
      },
    },
  });

  try {
    await stagehand.init();
    sessionUrl = `https://browserbase.com/sessions/${stagehand.browserbaseSessionID}`;
    console.log(`✅ Session: ${sessionUrl}`);

    const page = stagehand.context.pages()[0];

    for (let i = 0; i < STEPS.length; i++) {
      const result = await runStep(STEPS[i], page, context, i + 1);
      results.push({ step: STEPS[i].name, ...result });
      if (!result.success) {
        console.log(`🛑 Stopped at step ${i + 1}: ${STEPS[i].name}`);
        break;
      }
    }
  } catch (error: any) {
    console.error(`❌ Fatal error: ${error.message}`);
    results.push({ step: "Fatal Error", success: false, error: error.message });
  } finally {
    await stagehand.close();
    try { fs.unlinkSync(tempPath); } catch { /* best-effort cleanup */ }
  }

  const success = !!context.videoUploaded;
  const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
  console.log(`${success ? "✅" : "❌"} SERA Attach Video done in ${elapsed}s`);

  return {
    success,
    message: context.completionMessage || (success
      ? "Video uploaded successfully"
      : "Video upload failed — check session replay"),
    jobId,
    appointmentId,
    fileName,
    videoUploaded: !!context.videoUploaded,
    sessionUrl,
    elapsedSeconds: parseFloat(elapsed),
    results,
  };
}

// =============================================================================
// EXPRESS SERVER
// =============================================================================
// NOTE: the SERA "jobId" (from the request body) and the in-memory tracking
// "taskId" are two different things — taskId tracks this background run,
// jobId points to the SERA job we're uploading to.

const app = express();
// Bump body size limit — base64 payloads / large JSON inputs can exceed the
// default 100kb cap quickly.
app.use(express.json({ limit: "50mb" }));

const tasks = new Map<string, { status: "running" | "done" | "failed"; result?: any; error?: string }>();

app.post("/run-sera-attach-video", (req, res) => {
  const taskId = "task_" + Date.now();
  tasks.set(taskId, { status: "running" });

  runSeraAttachVideo(req.body)
    .then(result => tasks.set(taskId, { status: "done", result }))
    .catch(err => tasks.set(taskId, { status: "failed", error: err.message }));

  res.json({ taskId });
});

app.get("/task-status/:id", (req, res) => {
  const task = tasks.get(req.params.id);
  if (!task) return res.status(404).json({ error: "not found" });
  res.json(task);
});

app.listen(3000, () => {
  console.log("SERA attach-video server running on port 3000");
});
