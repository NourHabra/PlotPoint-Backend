const rateLimit = require('express-rate-limit');
const helmet = require('helmet');
const mongoSanitize = require('express-mongo-sanitize');
const xss = require('xss-clean');
const hpp = require('hpp');

// Rate limiting configuration
const createRateLimiter = (windowMs, max, message) => {
  return rateLimit({
    windowMs,
    max,
    message: { error: message },
    standardHeaders: true,
    legacyHeaders: false,
    // Use a secure key generator
    keyGenerator: (req) => {
      return req.ip || req.connection.remoteAddress;
    },
  });
};

// General API rate limiter
const apiLimiter = createRateLimiter(
  15 * 60 * 1000, // 15 minutes
  500, // max 500 requests per windowMs
  'Too many requests from this IP, please try again later.'
);

// Strict rate limiter for authentication endpoints
const authLimiter = createRateLimiter(
  15 * 60 * 1000, // 15 minutes
  20, // max 20 requests per windowMs
  'Too many login attempts from this IP, please try again after 15 minutes.'
);

// Strict rate limiter for registration
const registerLimiter = createRateLimiter(
  60 * 60 * 1000, // 1 hour
  10, // max 10 requests per hour
  'Too many account creation attempts, please try again later.'
);

// Upload rate limiter
const uploadLimiter = createRateLimiter(
  60 * 60 * 1000, // 1 hour
  200, // max 200 uploads per hour
  'Too many upload requests, please try again later.'
);

// Helmet configuration for security headers
const helmetConfig = helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      styleSrc: ["'self'", "'unsafe-inline'"],
      scriptSrc: ["'self'"],
      imgSrc: ["'self'", 'data:', 'https:'],
      connectSrc: ["'self'"],
      fontSrc: ["'self'"],
      objectSrc: ["'none'"],
      mediaSrc: ["'self'"],
      frameSrc: ["'none'"],
    },
  },
  crossOriginEmbedderPolicy: true,
  crossOriginOpenerPolicy: true,
  crossOriginResourcePolicy: { policy: "same-site" },
  dnsPrefetchControl: true,
  frameguard: { action: 'deny' },
  hidePoweredBy: true,
  hsts: {
    maxAge: 31536000,
    includeSubDomains: true,
    preload: true
  },
  ieNoOpen: true,
  noSniff: true,
  referrerPolicy: { policy: 'strict-origin-when-cross-origin' },
  xssFilter: true,
});

// CORS configuration
const getCorsOptions = () => {
  const allowedOrigins = process.env.ALLOWED_ORIGINS 
    ? process.env.ALLOWED_ORIGINS.split(',').map(o => o.trim())
    : ['http://localhost:3000']; // Default for development only

  return {
    origin: function (origin, callback) {
      // Allow requests with no origin (like server-to-server, curl, or same-origin)
      // This happens when requests come through Nginx or from server-side
      if (!origin) {
        return callback(null, true);
      }
      
      // Check if origin is in allowed list
      if (allowedOrigins.indexOf(origin) !== -1 || allowedOrigins.includes('*')) {
        callback(null, true);
      } else {
        callback(new Error('Not allowed by CORS'));
      }
    },
    credentials: true,
    optionsSuccessStatus: 200,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    maxAge: 86400, // 24 hours
  };
};

// Input sanitization middleware
const sanitizeInput = [
  mongoSanitize(), // Prevent NoSQL injection
  xss(), // Prevent XSS attacks
  hpp(), // Prevent HTTP Parameter Pollution
];

module.exports = {
  apiLimiter,
  authLimiter,
  registerLimiter,
  uploadLimiter,
  helmetConfig,
  getCorsOptions,
  sanitizeInput,
};

