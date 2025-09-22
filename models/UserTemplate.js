const mongoose = require("mongoose");

const snippetSchema = new mongoose.Schema(
	{
		id: { type: String, required: true },
		text: { type: String, required: true },
	},
	{ _id: false }
);

const variableSnippetsSchema = new mongoose.Schema(
	{
		variableId: { type: String, required: true },
		snippets: { type: [snippetSchema], default: [] },
	},
	{ _id: false }
);

const checklistItemSchema = new mongoose.Schema(
	{
		id: { type: String, required: true },
		label: { type: String, required: true },
		required: { type: Boolean, default: false },
		order: { type: Number, default: 0 },
	},
	{ _id: false }
);

const userTemplateSchema = new mongoose.Schema(
	{
		userId: { type: String, required: true, index: true },
		templateId: {
			type: mongoose.Schema.Types.ObjectId,
			ref: "Template",
			required: true,
			index: true,
		},
		variableTextTemplates: { type: [variableSnippetsSchema], default: [] },
		checklist: { type: [checklistItemSchema], default: [] },
	},
	{ timestamps: true }
);

userTemplateSchema.index({ userId: 1, templateId: 1 }, { unique: true });

module.exports = mongoose.model("UserTemplate", userTemplateSchema);
