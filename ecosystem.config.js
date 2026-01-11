module.exports = {
	apps: [
		{
			name: "plotpoint-backend",
			script: "./server.js",
			instances: "max", // Use all available CPU cores
			exec_mode: "cluster",
			autorestart: true,
			watch: false,
			max_memory_restart: "2G",
			env_production: {
				NODE_ENV: "production",
				PORT: 5000,
			},
			error_file: "./logs/err.log",
			out_file: "./logs/out.log",
			log_file: "./logs/combined.log",
			time: true,
			// Graceful shutdown
			kill_timeout: 5000,
			wait_ready: true,
			listen_timeout: 10000,
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
