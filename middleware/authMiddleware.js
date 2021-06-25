const jwt = require("jsonwebtoken");
const asyncHandler = require("express-async-handler");
const User = require("../models/users.js");

const protect = asyncHandler(async (req, res, next) => {
  let token;

  if (
    req.headers.authorization &&
    req.headers.authorization.startsWith("Bearer")
  ) {
    try {
      token = req.headers.authorization.split(" ")[1];

      const decoded = jwt.verify(token, process.env.JWT_SECRET);
      req.user = await User.findByPk(decoded.userId);
      next();
    } catch (error) {
      res.status(401).json("Not authorized, token failed, Logging you out");
    }
  }

  if (!token) {
    res.status(401).json("Not authorized, no token, Logging you out");
  }
});

const accountant = asyncHandler(async (req, res, next) => {
  if (req.user && req.user.role === "accountant") {
    next();
  } else {
    res.status(401).json("Not authorized as an accountant, Logging you out");
  }
});
const ordermanager = asyncHandler(async (req, res, next) => {
  if (req.user && req.user.role === "ordermanager") {
    next();
  } else {
    res.status(401).json("Not authorized as an ordermanager, Logging you out");
  }
});

exports.protect = protect;
exports.accountant = accountant;
exports.ordermanager = ordermanager;
