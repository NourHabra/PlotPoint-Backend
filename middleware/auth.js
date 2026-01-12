const jwt = require('jsonwebtoken');

// Validate JWT secret exists
const JWT_SECRET = process.env.JWT_SECRET;
if (!JWT_SECRET || JWT_SECRET === 'devsecret') {
  console.error('❌ CRITICAL: JWT_SECRET not set or using default value!');
  if (process.env.NODE_ENV === 'production') {
    console.error('❌ CRITICAL: Cannot start in production without a secure JWT_SECRET!');
    process.exit(1);
  } else {
    console.warn('⚠️  WARNING: Using insecure JWT secret in development mode');
  }
}

// Verify JWT token middleware
const verifyToken = (req, res, next) => {
  try {
    const authHeader = req.headers.authorization;
    
    if (!authHeader) {
      return res.status(401).json({ error: 'No authorization header provided' });
    }

    const parts = authHeader.split(' ');
    
    if (parts.length !== 2 || parts[0] !== 'Bearer') {
      return res.status(401).json({ error: 'Invalid authorization header format. Use: Bearer <token>' });
    }

    const token = parts[1];
    
    if (!token) {
      return res.status(401).json({ error: 'No token provided' });
    }

    // Verify token
    const decoded = jwt.verify(token, JWT_SECRET);
    
    // Attach user info to request
    req.user = {
      id: decoded.userId || decoded.id,
      email: decoded.email,
      role: decoded.role,
    };
    
    next();
  } catch (error) {
    if (error.name === 'JsonWebTokenError') {
      return res.status(401).json({ error: 'Invalid token' });
    }
    if (error.name === 'TokenExpiredError') {
      return res.status(401).json({ error: 'Token expired' });
    }
    return res.status(500).json({ error: 'Token verification failed' });
  }
};

// Verify admin role middleware
const verifyAdmin = (req, res, next) => {
  if (!req.user) {
    return res.status(401).json({ error: 'Authentication required' });
  }
  
  if (req.user.role !== 'Admin') {
    return res.status(403).json({ error: 'Admin access required' });
  }
  
  next();
};

// Generate JWT token
const generateToken = (userId, email, role) => {
  if (!JWT_SECRET) {
    throw new Error('JWT_SECRET not configured');
  }
  
  return jwt.sign(
    { 
      sub: userId,        // Standard JWT claim for user ID
      userId,             // Keep for backwards compatibility
      email, 
      role,
      iat: Math.floor(Date.now() / 1000),
    },
    JWT_SECRET,
    { 
      expiresIn: '7d',
      issuer: 'plotpoint-system',
    }
  );
};

// Verify token without middleware (for manual checks)
const verifyTokenManual = (token) => {
  if (!JWT_SECRET) {
    throw new Error('JWT_SECRET not configured');
  }
  return jwt.verify(token, JWT_SECRET);
};

module.exports = {
  verifyToken,
  verifyAdmin,
  generateToken,
  verifyTokenManual,
  JWT_SECRET,
};

