module.exports = {
	apps: [
		{
			name: "plotpoint-backend",
			script: "./server.js",
			instances: 1, // One process only; server.js listens on one port (cluster would need app-level clustering)
			exec_mode: "fork",
			autorestart: true,
			watch: false,
			max_memory_restart: "2G",
		env_production: {
			NODE_ENV: "production",
			PORT: 5000,
			HOST: "127.0.0.1",
		},
			error_file: "./logs/err.log",
			out_file: "./logs/out.log",
			log_file: "./logs/combined.log",
			time: true,
			// Graceful shutdown
			kill_timeout: 5000,
			// NOTE: `wait_ready` requires the app to call `process.send("ready")`.
			// This backend does not, so leaving it enabled can cause restart loops.
			wait_ready: false,
			// Restart strategies
			min_uptime: "10s",
			max_restarts: 10,
			// Advanced features
			instance_var: "INSTANCE_ID",
			// Environment variables
			merge_logs: true,
			// Monitoring
			pmx: true,
			// Log rotation (requires pm2-logrotate module)
			// Install: pm2 install pm2-logrotate
		},
	],
};
