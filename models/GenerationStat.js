const mongoose = require("mongoose");

const generationStatSchema = new mongoose.Schema(
	{
		// Links
		reportId: { type: mongoose.Schema.Types.ObjectId, ref: "Report" },
		reportName: { type: String, default: "" },
		templateId: { type: mongoose.Schema.Types.ObjectId, ref: "Template" },
		templateName: { type: String, default: "" },
		// Auth/user
		userId: { type: String, default: "" },
		username: { type: String, default: "" },
		email: { type: String, default: "" },
		// Metrics
		output: { type: String, enum: ["docx", "pdf"], default: "docx" },
		durationMs: { type: Number, default: 0 },
		inlineImages: { type: Number, default: 0 },
		appendixItems: { type: Number, default: 0 },
		// When generation occurred (separate from createdAt)
		timestamp: { type: Date, default: Date.now },
	},
	{ timestamps: true }
);

generationStatSchema.index({ createdAt: -1 });

module.exports = mongoose.model("GenerationStat", generationStatSchema);
