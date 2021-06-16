const bcrypt = require("bcryptjs");
const asyncHandler = require("express-async-handler");

const jwt = require("jsonwebtoken");
const User = require("../models/users");

const addNewUser = asyncHandler(async (req, res, next) => {
  const { name, password, role } = req.body;

  let existingUser;
  try {
    existingUser = await User.findOne({ where: { name: name } });
  } catch (err) {
    return res.status(500).json("Signing up failed, please try again later.");
  }

  if (existingUser) {
    return res.status(422).json("User exists already, please login instead.");
  }

  let hashedPassword;
  try {
    hashedPassword = await bcrypt.hash(password, 12);
  } catch (err) {
    return res.status(500).json(err);
  }
  let createdUser;
  try {
    createdUser = await User.create({
      name,
      password: hashedPassword,
      role,
    });
  } catch (err) {
    if (err.errors[0].message) {
      return res.status(500).json(err.errors[0].message);
    } else {
      return res.status(500).json(err);
    }
  }

  return res.status(201).json({
    userId: createdUser.id,
    name: createdUser.name,
    role: createdUser.role,
  });
});

const login = asyncHandler(async (req, res, next) => {
  const { userName, password } = req.body;
  console.log(req.body);
  let existingUser;
  try {
    existingUser = await User.findOne({ where: { name: userName } });
  } catch (err) {
    return res.status(500).json("Logging in failed, please try again later.");
  }

  if (!existingUser) {
    return res.status(403).json("Invalid credentials, could not log you in.");
  }

  let isValidPassword = false;
  try {
    isValidPassword = await bcrypt.compare(password, existingUser.password);
  } catch (err) {
    return res
      .status(500)
      .json(
        "Could not log you in, please check your credentials and try again."
      );
  }

  if (!isValidPassword) {
    return res.status(403).json("Invalid credentials, could not log you in.");
  }

  let token;
  try {
    token = jwt.sign(
      {
        userId: existingUser.id,
        name: existingUser.name,
        role: existingUser.role,
      },
      process.env.JWT_SECRET,
      {
        expiresIn: "8h",
      }
    );
  } catch (err) {
    return res.status(500).json("Logging in failed, please try again later.");
  }
  return res.status(201).json({
    userId: existingUser.id,
    name: existingUser.name,
    role: existingUser.role,
    token: token,
  });
});

exports.addNewUser = addNewUser;
exports.login = login;
