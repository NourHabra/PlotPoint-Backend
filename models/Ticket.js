const mongoose = require("mongoose");

const ticketSchema = new mongoose.Schema(
	{
		title: { type: String, required: true },
		contactName: { type: String, required: true },
		contactEmail: { type: String, required: true },
		phoneCountryCode: { type: String, default: "" },
		phoneNumber: { type: String, default: "" },
		message: { type: String, required: true },
		status: {
			type: String,
			enum: ["Open", "Resolved", "Withdrawn"],
			default: "Open",
		},
		createdBy: { type: String, required: true },
		resolvedBy: { type: String, default: "" },
		resolvedAt: { type: Date },
		adminResponse: { type: String, default: "" },
		isArchived: { type: Boolean, default: false },
	},
	{ timestamps: true }
);

module.exports = mongoose.model("Ticket", ticketSchema);
