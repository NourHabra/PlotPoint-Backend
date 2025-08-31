const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const multer = require("multer");
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const { execFile } = require("child_process");
const mammoth = require("mammoth");
const pathLib = require("path");
const fetch = require("node-fetch");
const sharp = require("sharp");
require("dotenv").config();

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(cors());
app.use(express.json({ limit: "10mb" }));

// MongoDB Connection
mongoose
	.connect(
		process.env.MONGODB_URI || "mongodb://localhost:27017/template-db",
		{
			useNewUrlParser: true,
			useUnifiedTopology: true,
		}
	)
	.then(() => console.log("Connected to MongoDB"))
	.catch((err) => console.error("MongoDB connection error:", err));

// Import Template model
const Template = require("./models/Template");
const Report = require("./models/Report");
const Ticket = require("./models/Ticket");
const User = require("./models/User");
const jwt = require("jsonwebtoken");
const bcrypt = require("bcryptjs");

// Report status workflow (scalable)
const REPORT_STATUS_FLOW = [
	"Draft",
	"Initial Review",
	"Final Review",
	"Submitted",
];
function isValidNextStatus(current, next) {
	if (!next || current === next) return true; // no-op allowed
	const ci = REPORT_STATUS_FLOW.indexOf(current || "Draft");
	const ni = REPORT_STATUS_FLOW.indexOf(next);
	if (ni === -1) return false; // unknown target
	return ni === ci + 1; // only advance one step forward
}
// Resolve LibreOffice (soffice) path cross-platform
function resolveSofficePath() {
	if (process.env.LIBREOFFICE_PATH) return process.env.LIBREOFFICE_PATH;
	if (process.platform === "win32") {
		const candidates = [
			"C\\\\Program Files\\\\LibreOffice\\\\program\\\\soffice.com",
			"C\\\\Program Files\\\\LibreOffice\\\\program\\\\soffice.exe",
			"C\\\\Program Files (x86)\\\\LibreOffice\\\\program\\\\soffice.com",
			"C\\\\Program Files (x86)\\\\LibreOffice\\\\program\\\\soffice.exe",
		];
		for (const p of candidates) {
			try {
				if (fs.existsSync(p)) return p;
			} catch (_) {}
		}
	}
	return "soffice";
}

async function convertDocxToPdf(sofficePath, inputDocxPath, outputDir) {
	// Isolated profile to avoid lock/state issues
	const profileDir = path.join(outputDir, `lo-profile-${Date.now()}`);
	try {
		fs.mkdirSync(profileDir, { recursive: true });
	} catch (_) {}
	const profileUrl = `file:///${profileDir.replace(/\\/g, "/")}`;

	const makeArgs = (filter) => [
		"--headless",
		"--nocrashreport",
		"--nolockcheck",
		"--nodefault",
		"--nologo",
		"--norestore",
		`-env:UserInstallation=${profileUrl}`,
		"--convert-to",
		filter,
		"--outdir",
		outputDir,
		inputDocxPath,
	];

	const tryExec = (args) =>
		new Promise((resolve, reject) => {
			execFile(
				sofficePath,
				args,
				{ windowsHide: true },
				(err, stdout, stderr) => {
					if (err) {
						const detail =
							(stderr && stderr.toString()) ||
							(stdout && stdout.toString()) ||
							err.message;
						return reject(new Error(detail));
					}
					resolve();
				}
			);
		});

	try {
		await tryExec(makeArgs("pdf:writer_pdf_Export"));
	} catch (_) {
		await tryExec(makeArgs("pdf"));
	} finally {
		// Cleanup LibreOffice user profile directory
		try {
			fs.rmSync(profileDir, { recursive: true, force: true });
		} catch (_) {}
	}
}
// Storage for uploaded templates
const uploadsDir = path.join(__dirname, "uploads");
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir);
const templatePreviewsDir = path.join(uploadsDir, "template-previews");
const imagesDir = path.join(uploadsDir, "images");
try {
	fs.mkdirSync(templatePreviewsDir, { recursive: true });
} catch (_) {}
try {
	fs.mkdirSync(imagesDir, { recursive: true });
} catch (_) {}

const storage = multer.diskStorage({
	destination: function (req, file, cb) {
		cb(null, uploadsDir);
	},
	filename: function (req, file, cb) {
		const unique = Date.now() + "-" + Math.round(Math.random() * 1e9);
		cb(null, unique + path.extname(file.originalname));
	},
});
const upload = multer({ storage });

// Storage for image uploads (used for image variables)
const imageStorage = multer.diskStorage({
	destination: function (req, file, cb) {
		cb(null, imagesDir);
	},
	filename: function (req, file, cb) {
		const unique = Date.now() + "-" + Math.round(Math.random() * 1e9);
		cb(null, unique + path.extname(file.originalname));
	},
});
const uploadImage = multer({ storage: imageStorage });

// Dedicated storage for profile photos
const avatarsDir = path.join(__dirname, "uploads", "avatars");
try {
	fs.mkdirSync(avatarsDir, { recursive: true });
} catch (_) {}
const avatarStorage = multer.diskStorage({
	destination: function (req, file, cb) {
		cb(null, avatarsDir);
	},
	filename: function (req, file, cb) {
		const unique = Date.now() + "-" + Math.round(Math.random() * 1e9);
		cb(null, unique + path.extname(file.originalname));
	},
});
const uploadAvatar = multer({ storage: avatarStorage });

// Serve avatars as static assets
app.use("/uploads/avatars", express.static(avatarsDir));
// Serve template preview PDFs
app.use("/uploads/template-previews", express.static(templatePreviewsDir));
// Serve uploaded images for variables
app.use("/uploads/images", express.static(imagesDir));
// Helper: resolve an image buffer from provided value (data URL, absolute URL, or local uploads path)
async function resolveImageBuffer(provided) {
	try {
		if (!provided) return null;
		if (typeof provided !== "string") return null;
		if (provided.startsWith("data:")) {
			const base64 = provided.split(",")[1] || "";
			return Buffer.from(base64, "base64");
		}
		// Local uploads path (relative): /uploads/images/<file>
		const uploadPrefix = "/uploads/images/";
		if (
			provided.startsWith(uploadPrefix) ||
			provided.startsWith("uploads/images/")
		) {
			const fname = provided.replace(/^\/?uploads\/images\//, "");
			const full = path.join(imagesDir, fname);
			try {
				return fs.readFileSync(full);
			} catch (_) {
				return null;
			}
		}
		// Absolute URL (http/https)
		if (/^https?:\/\//i.test(provided)) {
			const resp = await fetch(provided);
			if (!resp.ok) return null;
			const ab = await resp.arrayBuffer();
			return Buffer.from(ab);
		}
		return null;
	} catch (_) {
		return null;
	}
}

function pickSharpFormatForTarget(targetPath) {
	const ext = String(path.extname(targetPath || "")).toLowerCase();
	if (ext === ".jpg" || ext === ".jpeg") return "jpeg";
	if (ext === ".png") return "png";
	if (ext === ".webp") return "webp";
	// default to png for unknown formats (EMF/WMF/TIFF will be re-encoded)
	return "png";
}

function resolveRelTargetToFullWordPath(relTarget) {
	// Normalize ../ paths and ensure under word/
	let t = String(relTarget || "").replace(/^\.{2}\//g, "");
	if (!t.startsWith("word/")) t = `word/${t}`;
	return t;
}

function findExtentForTargetInZip(zip, targetFullPath) {
	try {
		const allXmlPaths = Object.keys(zip.files || {}).filter(
			(k) =>
				k.startsWith("word/") &&
				k.endsWith(".xml") &&
				!k.includes("/_rels/")
		);
		for (const xmlPath of allXmlPaths) {
			const base = xmlPath.split("/").pop();
			const relsPath = `word/_rels/${base}.rels`;
			const xml = zip.file(xmlPath)?.asText() || "";
			const relsXml = zip.file(relsPath)?.asText() || "";
			const rels = {};
			let m;
			const reRel =
				/<Relationship[^>]*Id=\"([^\"]+)\"[^>]*Target=\"([^\"]+)\"/g;
			while ((m = reRel.exec(relsXml))) {
				rels[m[1]] = resolveRelTargetToFullWordPath(m[2]);
			}
			for (const [rid, full] of Object.entries(rels)) {
				if (full === targetFullPath) {
					const reBlip = new RegExp(
						`<a:blip[^>]*r:(?:embed|link)=\\\"${rid}\\\"[^>]*\\/>`,
						"g"
					);
					let blipMatch;
					while ((blipMatch = reBlip.exec(xml))) {
						const start = Math.max(0, blipMatch.index - 1500);
						const end = Math.min(
							xml.length,
							blipMatch.index + 1500
						);
						const slice = xml.slice(start, end);
						const mm =
							/<wp:extent[^>]*cx=\"([0-9]+)\"[^>]*cy=\"([0-9]+)\"/i.exec(
								slice
							);
						if (mm) {
							return {
								cx: Number(mm[1] || 0),
								cy: Number(mm[2] || 0),
							};
						}
					}
				}
			}
		}
	} catch (_) {}
	return null;
}

function computeSrcRectCover(pxW, pxH, imgW, imgH) {
	if (!pxW || !pxH || !imgW || !imgH) return null;
	const frameAR = pxW / pxH;
	const imgAR = imgW / imgH;
	let l = 0,
		r = 0,
		t = 0,
		b = 0;
	if (imgAR > frameAR) {
		// crop horizontally
		const targetW = frameAR * imgH;
		const crop = (imgW - targetW) / imgW; // proportion to remove
		const side = crop / 2;
		l = r = Math.max(0, Math.min(1, side));
	} else if (imgAR < frameAR) {
		// crop vertically
		const targetH = imgW / frameAR;
		const crop = (imgH - targetH) / imgH;
		const side = crop / 2;
		t = b = Math.max(0, Math.min(1, side));
	}
	const toPct = (v) => String(Math.round(v * 100000));
	return { l: toPct(l), r: toPct(r), t: toPct(t), b: toPct(b) };
}

function applySrcRectForRid(xml, rid, srcRect) {
	try {
		if (!srcRect) return xml;
		// Find the blipFill block that contains this rId
		const reBlip = new RegExp(
			`<a:blip[^>]*r:(?:embed|link)=\\\"${rid}\\\"[^>]*\\/>`
		);
		const blipMatch = reBlip.exec(xml);
		if (!blipMatch) return xml;
		// Find surrounding blipFill block
		const afterIdx = blipMatch.index + blipMatch[0].length;
		const beforeIdx = xml.lastIndexOf("<", blipMatch.index);
		// Search forward for closing of blipFill
		const reBlipFill = /<([a-zA-Z0-9]+:)?blipFill[^>]*>/g;
		reBlipFill.lastIndex = beforeIdx >= 0 ? beforeIdx : 0;
		let openIdx = -1;
		let openTag = null;
		let m;
		while ((m = reBlipFill.exec(xml))) {
			const tagStart = m.index;
			const tagEnd = reBlipFill.lastIndex;
			// ensure this open tag is before our blip
			if (tagEnd <= blipMatch.index + blipMatch[0].length) {
				openIdx = tagStart;
				openTag = m[0];
			} else {
				break;
			}
		}
		if (openIdx === -1) return xml;
		const prefixMatch = openTag.match(/^<([a-zA-Z0-9]+:)?blipFill/);
		const nsPrefix = prefixMatch && prefixMatch[1] ? prefixMatch[1] : "";
		const closeTag = `</${nsPrefix}blipFill>`;
		const closeIdx = xml.indexOf(closeTag, afterIdx);
		if (closeIdx === -1) return xml;
		const blockStart = openIdx;
		const blockEnd = closeIdx + closeTag.length;
		const block = xml.slice(blockStart, blockEnd);
		// Remove any existing stretch and srcRect
		let newBlock = block
			.replace(/<a:stretch>[\s\S]*?<\/a:stretch>/g, "")
			.replace(/<a:srcRect[^>]*\/>/g, "");
		// Insert srcRect before closing tag
		const insert = `<a:srcRect l="${srcRect.l}" t="${srcRect.t}" r="${srcRect.r}" b="${srcRect.b}"/>`;
		newBlock = newBlock.replace(closeTag, insert + closeTag);
		return xml.slice(0, blockStart) + newBlock + xml.slice(blockEnd);
	} catch (_) {
		return xml;
	}
}

function applyCoverCroppingForTarget(
	zip,
	targetFullPath,
	imgW,
	imgH,
	extentPxW,
	extentPxH
) {
	try {
		if (!extentPxW || !extentPxH || !imgW || !imgH) return;
		const srcRect = computeSrcRectCover(extentPxW, extentPxH, imgW, imgH);
		if (!srcRect) return;
		const allXmlPaths = Object.keys(zip.files || {}).filter(
			(k) =>
				k.startsWith("word/") &&
				k.endsWith(".xml") &&
				!k.includes("/_rels/")
		);
		for (const xmlPath of allXmlPaths) {
			const base = xmlPath.split("/").pop();
			const relsPath = `word/_rels/${base}.rels`;
			const xmlFile = zip.file(xmlPath);
			if (!xmlFile) continue;
			const xml = xmlFile.asText();
			const relsXml = zip.file(relsPath)?.asText() || "";
			const rels = {};
			let m;
			const reRel =
				/<Relationship[^>]*Id=\"([^\"]+)\"[^>]*Target=\"([^\"]+)\"/g;
			while ((m = reRel.exec(relsXml))) {
				rels[m[1]] = resolveRelTargetToFullWordPath(m[2]);
			}
			let updated = xml;
			for (const [rid, full] of Object.entries(rels)) {
				if (full === targetFullPath) {
					const next = applySrcRectForRid(updated, rid, srcRect);
					if (next !== updated) {
						updated = next;
						// do not break; in case multiple occurrences in same part
					}
				}
			}
			if (updated !== xml) {
				zip.file(xmlPath, updated);
			}
		}
	} catch (_) {}
}

// Routes
app.get("/", (req, res) => {
	res.json({ message: "Template API is running" });
});

function getAuthPayload(req) {
	const auth = req.headers.authorization || "";
	const token = auth.startsWith("Bearer ") ? auth.slice(7) : "";
	if (!token) return null;
	try {
		return jwt.verify(token, process.env.JWT_SECRET || "devsecret");
	} catch (_) {
		return null;
	}
}

// Auth routes
app.post(
	"/api/auth/register",
	uploadAvatar.single("avatar"),
	async (req, res) => {
		try {
			const { name, email, password, role = "User" } = req.body || {};
			if (!name || !email || !password)
				return res
					.status(400)
					.json({ message: "name, email, password required" });
			// After first user exists, only Admin can create accounts
			const totalUsers = await User.countDocuments();
			if (totalUsers > 0) {
				const auth = req.headers.authorization || "";
				const token = auth.startsWith("Bearer ") ? auth.slice(7) : "";
				try {
					const payload = jwt.verify(
						token,
						process.env.JWT_SECRET || "devsecret"
					);
					if (!payload || payload.role !== "Admin") {
						return res
							.status(403)
							.json({ message: "Admin privileges required" });
					}
				} catch (_) {
					return res
						.status(401)
						.json({ message: "Invalid or missing token" });
				}
			}
			const existing = await User.findOne({ email });
			if (existing)
				return res
					.status(409)
					.json({ message: "Email already in use" });
			const salt = await bcrypt.genSalt(10);
			const passwordHash = await bcrypt.hash(password, salt);
			const avatarPath = req.file
				? `/uploads/avatars/${req.file.filename}`
				: "";
			const user = new User({
				name,
				email,
				passwordHash,
				role: role === "Admin" ? "Admin" : "User",
				avatarPath,
			});
			const saved = await user.save();
			const token = jwt.sign(
				{
					sub: saved._id,
					role: saved.role,
					email: saved.email,
					name: saved.name,
				},
				process.env.JWT_SECRET || "devsecret",
				{ expiresIn: "7d" }
			);
			res.status(201).json({
				id: saved._id,
				name: saved.name,
				email: saved.email,
				role: saved.role,
				avatarUrl: saved.avatarPath,
				token,
			});
		} catch (error) {
			res.status(400).json({ message: error.message });
		}
	}
);

// Upload endpoint for image variables
app.post("/api/uploads/image", uploadImage.single("file"), async (req, res) => {
	try {
		if (!req.file)
			return res.status(400).json({ message: "file is required" });
		return res.status(201).json({
			filename: req.file.filename,
			url: `/uploads/images/${req.file.filename}`,
		});
	} catch (error) {
		return res.status(400).json({ message: error.message });
	}
});

app.post("/api/auth/login", async (req, res) => {
	try {
		const { email, password } = req.body || {};
		if (!email || !password)
			return res
				.status(400)
				.json({ message: "email and password required" });
		const user = await User.findOne({ email });
		if (!user)
			return res.status(401).json({ message: "Invalid credentials" });
		const ok = await bcrypt.compare(password, user.passwordHash);
		if (!ok)
			return res.status(401).json({ message: "Invalid credentials" });
		const token = jwt.sign(
			{
				sub: user._id,
				role: user.role,
				email: user.email,
				name: user.name,
			},
			process.env.JWT_SECRET || "devsecret",
			{ expiresIn: "7d" }
		);
		res.json({
			id: user._id,
			name: user.name,
			email: user.email,
			role: user.role,
			avatarUrl: user.avatarPath,
			token,
		});
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// List users (admin only)
app.get("/api/users", async (req, res) => {
	try {
		const auth = req.headers.authorization || "";
		const token = auth.startsWith("Bearer ") ? auth.slice(7) : "";
		const payload = token
			? (() => {
					try {
						return jwt.verify(
							token,
							process.env.JWT_SECRET || "devsecret"
						);
					} catch {
						return null;
					}
			  })()
			: null;
		if (!payload || payload.role !== "Admin")
			return res
				.status(403)
				.json({ message: "Admin privileges required" });
		const users = await User.find({}, { passwordHash: 0 }).sort({
			createdAt: -1,
		});
		res.json(users);
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// Current user profile
app.get("/api/users/me", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const user = await User.findById(payload.sub);
		if (!user) return res.status(404).json({ message: "User not found" });
		return res.json({
			id: user._id,
			name: user.name,
			email: user.email,
			role: user.role,
			avatarUrl: user.avatarPath,
			createdAt: user.createdAt,
		});
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// Shared generator used by template and report routes
async function generateFromTemplate(template, inputValues, output, res) {
	if (!template)
		return res.status(404).json({ message: "Template not found" });
	if (!template.sourceDocxPath)
		return res
			.status(400)
			.json({ message: "Template was not imported from DOCX" });

	// Load DOCX file
	const content = fs.readFileSync(template.sourceDocxPath, "binary");
	const zip = new PizZip(content);

	// Inject images for image variables before text rendering
	try {
		const mediaPrefix = "word/media/";
		for (const v of template.variables || []) {
			if (v.type !== "image") continue;
			const varKey = v.name;
			const provided = inputValues ? inputValues[varKey] : undefined;
			const imgBuffer = await resolveImageBuffer(provided);
			if (!imgBuffer) continue;
			let target =
				v.imageTarget && v.imageTarget.startsWith("word/")
					? v.imageTarget
					: null;
			if (!target) {
				const candidates = Object.keys(zip.files || {}).filter((k) =>
					k.startsWith(mediaPrefix)
				);
				if (candidates.length) target = candidates[0];
			}
			if (!target) continue;
			let finalBuf = imgBuffer;
			try {
				const fmt = pickSharpFormatForTarget(target);
				let extent =
					v.imageExtent && v.imageExtent.cx && v.imageExtent.cy
						? { cx: v.imageExtent.cx, cy: v.imageExtent.cy }
						: findExtentForTargetInZip(zip, target) || null;
				if (extent && extent.cx && extent.cy) {
					const pxW = Math.max(
						1,
						Math.round((extent.cx / 914400) * 96)
					);
					const pxH = Math.max(
						1,
						Math.round((extent.cy / 914400) * 96)
					);
					const meta = await sharp(imgBuffer).metadata();
					finalBuf = await sharp(imgBuffer)
						.resize(pxW, pxH, { fit: "cover", position: "center" })
						.toFormat(fmt)
						.toBuffer();
					if (meta && meta.width && meta.height) {
						applyCoverCroppingForTarget(
							zip,
							target,
							meta.width,
							meta.height,
							pxW,
							pxH
						);
					}
				} else {
					finalBuf = await sharp(imgBuffer).toFormat(fmt).toBuffer();
				}
			} catch (_) {}
			zip.file(target, finalBuf);
		}
	} catch (_) {}
	const doc = new Docxtemplater(zip, {
		paragraphLoop: true,
		linebreaks: true,
		delimiters: { start: "{{", end: "}}" },
		nullGetter: (part) => {
			if (part.module === "rawxml") return "";
			return part.tag ? `[${part.tag}]` : "";
		},
	});

	// Build final values map and evaluate calculateds; map KML aliases
	const finalValues = { ...(inputValues || {}) };
	for (const v of template.variables || []) {
		if (
			v.type === "kml" &&
			v.kmlField &&
			finalValues[v.kmlField] &&
			!finalValues[v.name]
		) {
			finalValues[v.name] = finalValues[v.kmlField];
		}
		if (v.type === "calculated" && v.expression) {
			try {
				// eslint-disable-next-line no-new-func
				const fn = new Function(
					...Object.keys(finalValues),
					`return (${v.expression});`
				);
				finalValues[v.name] = String(fn(...Object.values(finalValues)));
			} catch (e) {
				finalValues[v.name] = "";
			}
		}
	}

	// Convert empty strings to null so nullGetter is used
	Object.keys(finalValues).forEach((k) => {
		if (finalValues[k] === "") finalValues[k] = null;
	});
	try {
		doc.render(finalValues);
	} catch (error) {
		return res.status(400).json({
			message: "Template rendering failed",
			detail: error.message,
		});
	}

	const buf = doc.getZip().generate({ type: "nodebuffer" });
	const outDocx = path.join(uploadsDir, `out-${Date.now()}.docx`);
	fs.writeFileSync(outDocx, buf);

	if (output === "pdf") {
		const sofficePrimary = resolveSofficePath();
		// Preflight: verify soffice is reachable (supports PATH)
		try {
			await new Promise((resolve, reject) => {
				execFile(sofficePrimary, ["--headless", "--version"], (err) =>
					err ? reject(err) : resolve()
				);
			});
		} catch (err) {
			try {
				fs.rmSync(outDocx, { force: true });
			} catch (_) {}
			return res.status(500).json({
				message:
					"LibreOffice not found or not accessible. Install LibreOffice or set LIBREOFFICE_PATH to soffice(.com)",
				detail: err && err.message ? err.message : String(err),
			});
		}
		const outDir = uploadsDir;
		// Try conversion with primary path first
		let convertError = null;
		try {
			await convertDocxToPdf(sofficePrimary, outDocx, outDir);
		} catch (e1) {
			convertError = e1;
			// On Windows, retry with explicit soffice.com if available
			if (process.platform === "win32") {
				const alt = [
					"C\\\\Program Files\\\\LibreOffice\\\\program\\\\soffice.com",
					"C\\\\Program Files (x86)\\\\LibreOffice\\\\program\\\\soffice.com",
				].find((p) => {
					try {
						return fs.existsSync(p);
					} catch (_) {
						return false;
					}
				});
				if (alt) {
					try {
						await convertDocxToPdf(alt, outDocx, outDir);
						convertError = null;
					} catch (e2) {
						convertError = e2;
					}
				}
			}
		}
		if (convertError) {
			// Cleanup any created DOCX on conversion failure
			try {
				fs.rmSync(outDocx, { force: true });
			} catch (_) {}
			return res.status(500).json({
				message: "PDF conversion failed",
				detail: convertError.message,
			});
		}
		const outPdf = outDocx.replace(/\.docx$/, ".pdf");
		res.setHeader("Content-Type", "application/pdf");
		res.setHeader("Content-Disposition", `attachment; filename=report.pdf`);
		const pdfStream = fs.createReadStream(outPdf);
		pdfStream.pipe(res);
		const cleanupPdfAndDocx = () => {
			try {
				fs.rmSync(outPdf, { force: true });
			} catch (_) {}
			try {
				fs.rmSync(outDocx, { force: true });
			} catch (_) {}
		};
		pdfStream.on("error", cleanupPdfAndDocx);
		pdfStream.on("close", cleanupPdfAndDocx);
		res.on("finish", cleanupPdfAndDocx);
		res.on("close", cleanupPdfAndDocx);
		res.on("error", cleanupPdfAndDocx);
		return;
	}

	res.setHeader(
		"Content-Type",
		"application/vnd.openxmlformats-officedocument.wordprocessingml.document"
	);
	res.setHeader("Content-Disposition", `attachment; filename=report.docx`);
	const docxStream = fs.createReadStream(outDocx);
	docxStream.pipe(res);
	const cleanupDocx = () => {
		try {
			fs.rmSync(outDocx, { force: true });
		} catch (_) {}
	};
	docxStream.on("error", cleanupDocx);
	docxStream.on("close", cleanupDocx);
	res.on("finish", cleanupDocx);
	res.on("close", cleanupDocx);
	res.on("error", cleanupDocx);
}

// Get all templates (including inactive ones)
app.get("/api/templates", async (req, res) => {
	try {
		const { includeInactive = "false" } = req.query;
		const query = includeInactive === "true" ? {} : { isActive: true };
		const templates = await Template.find(query).sort({ updatedAt: -1 });
		res.json(templates);
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// Get template by ID
app.get("/api/templates/:id", async (req, res) => {
	try {
		const template = await Template.findById(req.params.id);
		if (!template) {
			return res.status(404).json({ message: "Template not found" });
		}
		res.json(template);
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// Create new template
app.post("/api/templates", async (req, res) => {
	try {
		const template = new Template(req.body);
		const savedTemplate = await template.save();
		res.status(201).json(savedTemplate);
	} catch (error) {
		res.status(400).json({ message: error.message });
	}
});

// Import DOCX template with variables metadata
app.post(
	"/api/templates/import-docx",
	upload.single("file"),
	async (req, res) => {
		try {
			const {
				name,
				description = "",
				variables,
				requiresKml = "false",
				variableGroups,
			} = req.body;
			if (!name)
				return res.status(400).json({ message: "name is required" });
			if (!req.file)
				return res.status(400).json({ message: "file is required" });

			const parsedVariables = variables ? JSON.parse(variables) : [];
			const parsedGroups = variableGroups
				? JSON.parse(variableGroups)
				: [];

			// Attempt to inject tokens {{name}} into a working copy of the DOCX
			let finalDocxPath = req.file.path;
			let tokenizedXml = ""; // capture xml used for tokenization for later verification
			try {
				const bin = fs.readFileSync(req.file.path, "binary");
				const zip = new PizZip(bin);
				const docXmlPath = "word/document.xml";
				const file = zip.file(docXmlPath);
				if (file) {
					let xml = file.asText();
					const escapeRegExp = (s) =>
						s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
					const nbspToSpace = (s) => s.replace(/\u00A0/g, " ");

					// Build segments of <w:t> content and mapping
					const reNode = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
					const segments = [];
					let m;
					while ((m = reNode.exec(xml))) {
						const full = m[0];
						const content = m[1];
						const fullStart = m.index;
						const openIdx = full.indexOf(">");
						const contentStart = fullStart + openIdx + 1;
						const contentEnd = contentStart + content.length;
						segments.push({
							contentStart,
							contentEnd,
							text: content,
						});
					}

					const buildJoined = () => {
						let joined = "";
						const map = [];
						for (let i = 0; i < segments.length; i++) {
							const t = segments[i].text;
							const norm = nbspToSpace(t);
							for (let j = 0; j < norm.length; j++) {
								joined += norm[j];
								map.push({ segIdx: i, offset: j });
							}
						}
						return { joined, map };
					};

					const applyPatches = (xmlInput, patches) => {
						patches.sort((a, b) => b.start - a.start);
						let out = xmlInput;
						for (const p of patches) {
							out =
								out.slice(0, p.start) +
								p.replacement +
								out.slice(p.end);
						}
						return out;
					};

					const injectToken = (varName, sourceText) => {
						if (!varName || !sourceText) return false;
						const token = `{{${varName}}}`;
						let replacedAny = false;
						while (true) {
							const { joined, map } = buildJoined();
							const hay = joined;
							const needle = nbspToSpace(sourceText);
							const startIdx = hay.indexOf(needle);
							if (startIdx === -1) break;
							const endIdx = startIdx + needle.length - 1;
							const startMap = map[startIdx];
							const endMap = map[endIdx];
							if (!startMap || !endMap) break;

							const patches = [];
							const startSeg = segments[startMap.segIdx];
							const endSeg = segments[endMap.segIdx];
							// Patch end segment first if different: remove leading covered part
							if (startMap.segIdx !== endMap.segIdx) {
								patches.push({
									start: endSeg.contentStart,
									end:
										endSeg.contentStart + endMap.offset + 1,
									replacement: "",
								});
								// Remove full content of intermediate segments
								for (
									let i = startMap.segIdx + 1;
									i <= endMap.segIdx - 1;
									i++
								) {
									const seg = segments[i];
									patches.push({
										start: seg.contentStart,
										end: seg.contentEnd,
										replacement: "",
									});
								}
								// Replace tail of start segment with token
								patches.push({
									start:
										startSeg.contentStart + startMap.offset,
									end: startSeg.contentEnd,
									replacement: token,
								});
							} else {
								// Single segment replacement
								patches.push({
									start:
										startSeg.contentStart + startMap.offset,
									end:
										startSeg.contentStart +
										endMap.offset +
										1,
									replacement: token,
								});
							}

							xml = applyPatches(xml, patches);
							// Update segments positions after patches: re-parse for next iteration
							segments.length = 0;
							reNode.lastIndex = 0;
							while ((m = reNode.exec(xml))) {
								const full = m[0];
								const content = m[1];
								const fullStart = m.index;
								const openIdx = full.indexOf(">");
								const contentStart = fullStart + openIdx + 1;
								const contentEnd =
									contentStart + content.length;
								segments.push({
									contentStart,
									contentEnd,
									text: content,
								});
							}
							replacedAny = true;
						}
						return replacedAny;
					};

					for (const v of parsedVariables) {
						injectToken(v.name, v.sourceText);
					}

					// Normalize any whitespace inside tokens
					xml = xml
						.replace(/\{\{\s+/g, "{{")
						.replace(/\s+\}\}/g, "}}");
					tokenizedXml = xml;
					zip.file(docXmlPath, xml);
					const out = zip.generate({ type: "nodebuffer" });
					const outPath = path.join(
						uploadsDir,
						`tokenized-${Date.now()}.docx`
					);
					fs.writeFileSync(outPath, out);
					finalDocxPath = outPath;
					// Remove the original uploaded DOCX to conserve storage since we keep the tokenized copy
					try {
						fs.rmSync(req.file.path, { force: true });
					} catch (_) {}
				}
			} catch (e) {
				console.warn(
					"Token injection failed; storing original DOCX",
					e.message
				);
			}

			// Mark variables that were successfully tokenized (very simple heuristic)
			const verifyTokenized = (xml, v) =>
				v && v.name && xml.includes(`{{${v.name}}}`);

			// Determine xml to verify: prefer tokenizedXml, else reload from finalDocxPath
			let xmlForVerify = tokenizedXml;
			if (!xmlForVerify) {
				try {
					const bin2 = fs.readFileSync(finalDocxPath, "binary");
					const zip2 = new PizZip(bin2);
					const f2 = zip2.file("word/document.xml");
					if (f2) xmlForVerify = f2.asText();
				} catch (_) {
					xmlForVerify = "";
				}
			}

			const template = new Template({
				name,
				description: description || "Imported Word template",
				requiresKml: String(requiresKml) === "true",
				createdBy: "system",
				sections: [],
				sourceDocxPath: finalDocxPath,
				variables: parsedVariables.map((v) => ({
					...v,
					tokenized: verifyTokenized(xmlForVerify, v),
				})),
				variableGroups: Array.isArray(parsedGroups) ? parsedGroups : [],
			});
			// Build and store an unfilled PDF preview for fast stage-1 display
			try {
				const sofficePrimary = resolveSofficePath();
				await convertDocxToPdf(
					sofficePrimary,
					finalDocxPath,
					templatePreviewsDir
				);
				const baseName = path.basename(
					finalDocxPath,
					path.extname(finalDocxPath)
				);
				const previewPdf = path.join(
					templatePreviewsDir,
					`${baseName}.pdf`
				);
				if (fs.existsSync(previewPdf)) {
					template.previewPdfPath = `/uploads/template-previews/${path.basename(
						previewPdf
					)}`;
				}
			} catch (e) {
				console.warn(
					"Failed to build template preview PDF:",
					e.message
				);
			}
			const saved = await template.save();
			res.status(201).json(saved);
		} catch (error) {
			console.error(error);
			res.status(400).json({ message: error.message });
		}
	}
);

// Generate filled DOCX and optionally PDF
app.post("/api/templates/:id/generate", async (req, res) => {
	try {
		const { id } = req.params;
		const { values, output = "docx" } = req.body; // values: { [name]: value }
		const template = await Template.findById(id);
		return generateFromTemplate(template, values, output, res);
	} catch (error) {
		console.error(error);
		res.status(500).json({ message: error.message });
	}
});

// Reports CRUD
app.post("/api/reports", async (req, res) => {
	try {
		const {
			templateId,
			name = "",
			title = "",
			values = {},
			kmlData,
		} = req.body || {};
		if (!templateId)
			return res.status(400).json({ message: "templateId is required" });
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const tpl = await Template.findById(templateId);
		if (!tpl)
			return res.status(404).json({ message: "Template not found" });

		// Merge KML data into values for any KML variables
		const mergedValues = { ...(values || {}) };
		if (kmlData && typeof kmlData === "object") {
			for (const v of tpl.variables || []) {
				if (
					v.type === "kml" &&
					v.kmlField &&
					Object.prototype.hasOwnProperty.call(kmlData, v.kmlField)
				) {
					const key = v.name || v.kmlField;
					mergedValues[key] = String(kmlData[v.kmlField] ?? "");
				}
			}
		}
		const report = new Report({
			templateId,
			templateName: tpl.name,
			name,
			title,
			values: mergedValues,
			...(kmlData !== undefined && { kmlData }),
			status: "Draft",
			createdBy: String(payload.sub || payload.email || "unknown"),
		});
		const saved = await report.save();
		res.status(201).json(saved);
	} catch (error) {
		res.status(400).json({ message: error.message });
	}
});

app.get("/api/reports", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const reports = await Report.find({
			createdBy: String(payload.sub || payload.email),
		}).sort({ updatedAt: -1 });
		res.json(reports);
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// Dashboard summary for current user
app.get("/api/dashboard/reports/summary", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const isAdmin = payload.role === "Admin";
		const userId = String(payload.sub || payload.email);
		const baseFilter = isAdmin ? {} : { createdBy: userId };
		const [draft, initial, final, submitted] = await Promise.all([
			Report.countDocuments({ ...baseFilter, status: "Draft" }),
			Report.countDocuments({ ...baseFilter, status: "Initial Review" }),
			Report.countDocuments({ ...baseFilter, status: "Final Review" }),
			Report.countDocuments({ ...baseFilter, status: "Submitted" }),
		]);
		res.json({
			draft,
			underReview: initial + final,
			submitted,
			total: draft + initial + final + submitted,
		});
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// Latest N reports for current user
app.get("/api/dashboard/reports/latest", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const isAdmin = payload.role === "Admin";
		const userId = String(payload.sub || payload.email);
		const limit = Math.max(1, Math.min(50, Number(req.query.limit) || 10));
		const filter = isAdmin ? {} : { createdBy: userId };
		const reports = await Report.find(filter)
			.sort({ updatedAt: -1 })
			.limit(limit);
		res.json(reports);
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// Support Tickets
// Create ticket (user only)
app.post("/api/tickets", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const {
			title,
			contactName,
			contactEmail,
			phoneCountryCode = "",
			phoneNumber = "",
			message,
		} = req.body || {};
		if (!title || !contactName || !contactEmail || !message) {
			return res.status(400).json({
				message:
					"title, contactName, contactEmail and message are required",
			});
		}
		const ticket = new Ticket({
			title,
			contactName,
			contactEmail,
			phoneCountryCode,
			phoneNumber,
			message,
			status: "Open",
			createdBy: String(payload.sub || payload.email),
		});
		const saved = await ticket.save();
		res.status(201).json(saved);
	} catch (error) {
		res.status(400).json({ message: error.message });
	}
});

// List current user's tickets
app.get("/api/tickets", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const tickets = await Ticket.find({
			createdBy: String(payload.sub || payload.email),
			status: { $ne: "Withdrawn" },
		}).sort({ updatedAt: -1 });
		res.json(tickets);
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// Admin: list all tickets
app.get("/api/tickets/all", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		if (payload.role !== "Admin")
			return res
				.status(403)
				.json({ message: "Admin privileges required" });
		const tickets = await Ticket.find({}).sort({ updatedAt: -1 }).lean();
		// Enrich with creator and resolver details (name/email/role/avatar)
		const isObjectId = (v) =>
			typeof v === "string" && /^[0-9a-fA-F]{24}$/.test(v);
		const byIdKeys = tickets
			.map((t) => t.createdBy)
			.filter((v) => isObjectId(v));
		const byEmailKeys = tickets
			.map((t) => t.createdBy)
			.filter((v) => typeof v === "string" && !isObjectId(v));
		const resByIdKeys = tickets
			.map((t) => t.resolvedBy)
			.filter((v) => isObjectId(v));
		const resByEmailKeys = tickets
			.map((t) => t.resolvedBy)
			.filter((v) => typeof v === "string" && v && !isObjectId(v));
		const unique = (arr) => Array.from(new Set(arr));
		const [usersById, usersByEmail] = await Promise.all([
			User.find(
				{ _id: { $in: unique([...byIdKeys, ...resByIdKeys]) } },
				{ passwordHash: 0 }
			).lean(),
			User.find(
				{ email: { $in: unique([...byEmailKeys, ...resByEmailKeys]) } },
				{ passwordHash: 0 }
			).lean(),
		]);
		const idMap = new Map(usersById.map((u) => [String(u._id), u]));
		const emailMap = new Map(usersByEmail.map((u) => [String(u.email), u]));
		const enriched = tickets.map((t) => {
			const creatorUser = isObjectId(t.createdBy)
				? idMap.get(String(t.createdBy))
				: emailMap.get(String(t.createdBy));
			const resolverUser = t.resolvedBy
				? isObjectId(t.resolvedBy)
					? idMap.get(String(t.resolvedBy))
					: emailMap.get(String(t.resolvedBy))
				: null;
			return {
				...t,
				creator: creatorUser
					? {
							id: String(creatorUser._id),
							name: creatorUser.name,
							email: creatorUser.email,
							role: creatorUser.role,
							avatarUrl: creatorUser.avatarPath || "",
					  }
					: null,
				resolver: resolverUser
					? {
							id: String(resolverUser._id),
							name: resolverUser.name,
							email: resolverUser.email,
							role: resolverUser.role,
							avatarUrl: resolverUser.avatarPath || "",
					  }
					: null,
			};
		});
		res.json(enriched);
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// Admin: quick check if any open tickets exist
app.get("/api/tickets/has-open", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		if (payload.role !== "Admin")
			return res
				.status(403)
				.json({ message: "Admin privileges required" });
		const count = await Ticket.countDocuments({ status: "Open" });
		return res.json({ hasOpen: count > 0, count });
	} catch (error) {
		return res.status(500).json({ message: error.message });
	}
});

// Admin: update ticket (resolve, respond)
app.put("/api/tickets/:id", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const ticket = await Ticket.findById(req.params.id);
		if (!ticket)
			return res.status(404).json({ message: "Ticket not found" });
		if (payload.role !== "Admin")
			return res
				.status(403)
				.json({ message: "Admin privileges required" });
		const { status, adminResponse } = req.body || {};
		const updates = {};
		if (typeof adminResponse === "string")
			updates.adminResponse = adminResponse;
		if (status === "Resolved" && ticket.status !== "Resolved") {
			updates.status = "Resolved";
			updates.resolvedBy = String(payload.sub || payload.email);
			updates.resolvedAt = new Date();
		}
		if (status === "Open" && ticket.status === "Resolved") {
			updates.status = "Open";
			updates.resolvedBy = "";
			updates.resolvedAt = undefined;
		}
		const updated = await Ticket.findByIdAndUpdate(req.params.id, updates, {
			new: true,
			runValidators: true,
		});
		res.json(updated);
	} catch (error) {
		res.status(400).json({ message: error.message });
	}
});

// User: withdraw own ticket
app.post("/api/tickets/:id/withdraw", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const ticket = await Ticket.findById(req.params.id);
		if (!ticket)
			return res.status(404).json({ message: "Ticket not found" });
		if (String(ticket.createdBy) !== String(payload.sub || payload.email)) {
			return res.status(403).json({ message: "Forbidden" });
		}
		if (ticket.status === "Resolved") {
			return res
				.status(400)
				.json({ message: "Resolved tickets cannot be withdrawn" });
		}
		if (ticket.status === "Withdrawn") return res.json(ticket);
		const updated = await Ticket.findByIdAndUpdate(
			req.params.id,
			{ status: "Withdrawn" },
			{ new: true }
		);
		return res.json(updated);
	} catch (error) {
		return res.status(400).json({ message: error.message });
	}
});

// Admin: get all reports in the system
app.get("/api/reports/all", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		if (payload.role !== "Admin")
			return res
				.status(403)
				.json({ message: "Admin privileges required" });
		const reports = await Report.find({}).sort({ updatedAt: -1 });
		res.json(reports);
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

app.get("/api/reports/:id", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const report = await Report.findById(req.params.id);
		if (!report)
			return res.status(404).json({ message: "Report not found" });
		if (String(report.createdBy) !== String(payload.sub || payload.email))
			return res.status(403).json({ message: "Forbidden" });
		res.json(report);
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

app.put("/api/reports/:id", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const existing = await Report.findById(req.params.id);
		if (!existing)
			return res.status(404).json({ message: "Report not found" });
		if (String(existing.createdBy) !== String(payload.sub || payload.email))
			return res.status(403).json({ message: "Forbidden" });
		const { name, title, values, status, kmlData } = req.body || {};
		if (
			status !== undefined &&
			!isValidNextStatus(existing.status || "Draft", status)
		) {
			return res
				.status(400)
				.json({ message: "Invalid status transition" });
		}
		// If kmlData present, merge into values using template mapping
		let finalValues =
			values !== undefined ? { ...(values || {}) } : undefined;
		if (kmlData && typeof kmlData === "object") {
			try {
				const tpl = await Template.findById(existing.templateId);
				if (tpl) {
					finalValues = finalValues || { ...(existing.values || {}) };
					for (const v of tpl.variables || []) {
						if (
							v.type === "kml" &&
							v.kmlField &&
							Object.prototype.hasOwnProperty.call(
								kmlData,
								v.kmlField
							)
						) {
							const key = v.name || v.kmlField;
							finalValues[key] = String(
								kmlData[v.kmlField] ?? ""
							);
						}
					}
				}
			} catch (_) {}
		}
		const updated = await Report.findByIdAndUpdate(
			req.params.id,
			{
				...(name !== undefined && { name }),
				...(title !== undefined && { title }),
				...(finalValues !== undefined && { values: finalValues }),
				...(status !== undefined && { status }),
				...(kmlData !== undefined && { kmlData }),
			},
			{ new: true, runValidators: true }
		);
		if (!updated)
			return res.status(404).json({ message: "Report not found" });
		res.json(updated);
	} catch (error) {
		res.status(400).json({ message: error.message });
	}
});

// Generate from Report
app.post("/api/reports/:id/generate", async (req, res) => {
	try {
		const { id } = req.params;
		const { output = "docx" } = req.body || {};
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const report = await Report.findById(id);
		if (!report)
			return res.status(404).json({ message: "Report not found" });
		if (String(report.createdBy) !== String(payload.sub || payload.email))
			return res.status(403).json({ message: "Forbidden" });
		// Enforce status before download: must be Final Review or Submitted
		const allowedToDownload = ["Final Review", "Submitted"].includes(
			report.status || "Draft"
		);
		if (!allowedToDownload) {
			return res.status(400).json({
				message: "Report must be in 'Final Review' before download",
			});
		}
		const template = await Template.findById(report.templateId);
		// Intercept finish to auto-mark Submitted if not already
		const originalEnd = res.end.bind(res);
		let responded = false;
		res.end = function (...args) {
			if (!responded) {
				responded = true;
				try {
					if (report.status !== "Submitted") {
						Report.findByIdAndUpdate(id, {
							status: "Submitted",
						}).catch(() => {});
					}
				} catch (_) {}
			}
			return originalEnd(...args);
		};
		return generateFromTemplate(template, report.values, output, res);
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// Delete report (owner only) if not Submitted
app.delete("/api/reports/:id", async (req, res) => {
	try {
		const payload = getAuthPayload(req);
		if (!payload) return res.status(401).json({ message: "Unauthorized" });
		const report = await Report.findById(req.params.id);
		if (!report)
			return res.status(404).json({ message: "Report not found" });
		if (String(report.createdBy) !== String(payload.sub || payload.email)) {
			return res.status(403).json({ message: "Forbidden" });
		}
		if ((report.status || "Draft") === "Submitted") {
			return res
				.status(400)
				.json({ message: "Submitted reports cannot be deleted" });
		}
		await Report.findByIdAndDelete(req.params.id);
		return res.json({ message: "Report deleted" });
	} catch (error) {
		return res.status(500).json({ message: error.message });
	}
});

// Generate WYSIWYG HTML preview from filled DOCX
app.post("/api/templates/:id/preview-html", async (req, res) => {
	try {
		const { id } = req.params;
		const { values } = req.body;
		const template = await Template.findById(id);
		if (!template)
			return res.status(404).json({ message: "Template not found" });
		if (!template.sourceDocxPath)
			return res
				.status(400)
				.json({ message: "Template was not imported from DOCX" });

		// Render a temporary DOCX with values
		const content = fs.readFileSync(template.sourceDocxPath, "binary");
		const zip = new PizZip(content);
		const doc = new Docxtemplater(zip, {
			paragraphLoop: true,
			linebreaks: true,
			delimiters: { start: "{{", end: "}}" },
			nullGetter: (part) => {
				if (part.module === "rawxml") return "";
				return part.tag ? `[${part.tag}]` : "";
			},
		});

		// Inject images similar to DOCX/PDF generation (cover fit, center crop)
		try {
			const mediaPrefix = "word/media/";
			for (const v of template.variables || []) {
				if (v.type !== "image") continue;
				const varKey = v.name;
				const provided = values ? values[varKey] : undefined;
				const imgBuffer = await resolveImageBuffer(provided);
				if (!imgBuffer) continue;
				let target =
					v.imageTarget && v.imageTarget.startsWith("word/")
						? v.imageTarget
						: null;
				if (!target) {
					const candidates = Object.keys(zip.files || {}).filter(
						(k) => k.startsWith(mediaPrefix)
					);
					if (candidates.length) target = candidates[0];
				}
				if (!target) continue;
				let finalBuf = imgBuffer;
				try {
					const fmt = pickSharpFormatForTarget(target);
					if (v.imageExtent && v.imageExtent.cx && v.imageExtent.cy) {
						const pxW = Math.max(
							1,
							Math.round((v.imageExtent.cx / 914400) * 96)
						);
						const pxH = Math.max(
							1,
							Math.round((v.imageExtent.cy / 914400) * 96)
						);
						const meta = await sharp(imgBuffer).metadata();
						finalBuf = await sharp(imgBuffer)
							.resize(pxW, pxH, {
								fit: "cover",
								position: "center",
							})
							.toFormat(fmt)
							.toBuffer();
						if (meta && meta.width && meta.height) {
							applyCoverCroppingForTarget(
								zip,
								target,
								meta.width,
								meta.height,
								pxW,
								pxH
							);
						}
					} else {
						finalBuf = await sharp(imgBuffer)
							.toFormat(fmt)
							.toBuffer();
					}
				} catch (_) {}
				zip.file(target, finalBuf);
			}
		} catch (_) {}

		const finalValues = { ...(values || {}) };
		for (const v of template.variables || []) {
			if (
				v.type === "kml" &&
				v.kmlField &&
				finalValues[v.kmlField] &&
				!finalValues[v.name]
			) {
				finalValues[v.name] = finalValues[v.kmlField];
			}
			if (v.type === "calculated" && v.expression) {
				try {
					// eslint-disable-next-line no-new-func
					const fn = new Function(
						...Object.keys(finalValues),
						`return (${v.expression});`
					);
					finalValues[v.name] = String(
						fn(...Object.values(finalValues))
					);
				} catch (_) {
					finalValues[v.name] = "";
				}
			}
		}

		// Convert empty to null to reveal placeholders in preview
		Object.keys(finalValues).forEach((k) => {
			if (finalValues[k] === "") finalValues[k] = null;
		});
		try {
			doc.render(finalValues);
		} catch (error) {
			return res.status(400).json({
				message: "Template rendering failed",
				detail: error.message,
			});
		}

		const filledBuf = doc.getZip().generate({ type: "nodebuffer" });
		// Convert to HTML for WYSIWYG-ish preview in-memory without writing files
		const { value: html } = await mammoth.convertToHtml({
			buffer: filledBuf,
		});
		res.setHeader("Content-Type", "application/json");
		res.send({ html });
	} catch (error) {
		console.error(error);
		res.status(500).json({ message: error.message });
	}
});

// Update template
app.put("/api/templates/:id", async (req, res) => {
	try {
		const template = await Template.findByIdAndUpdate(
			req.params.id,
			req.body,
			{ new: true, runValidators: true }
		);
		if (!template) {
			return res.status(404).json({ message: "Template not found" });
		}
		res.json(template);
	} catch (error) {
		res.status(400).json({ message: error.message });
	}
});

// Delete template (soft delete)
app.delete("/api/templates/:id", async (req, res) => {
	try {
		const template = await Template.findByIdAndUpdate(
			req.params.id,
			{ isActive: false },
			{ new: true }
		);
		if (!template) {
			return res.status(404).json({ message: "Template not found" });
		}
		res.json({ message: "Template deleted successfully" });
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// Reactivate template
app.patch("/api/templates/:id/reactivate", async (req, res) => {
	try {
		const template = await Template.findByIdAndUpdate(
			req.params.id,
			{ isActive: true },
			{ new: true }
		);
		if (!template) {
			return res.status(404).json({ message: "Template not found" });
		}
		res.json({ message: "Template reactivated successfully", template });
	} catch (error) {
		res.status(500).json({ message: error.message });
	}
});

// Start server
app.listen(PORT, () => {
	console.log(`Server is running on port ${PORT}`);
});
