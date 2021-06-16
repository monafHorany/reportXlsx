const express = require("express");
const router = express.Router();
const userController = require("../controller/user");
const {
  protect,
  admin,
  adminANDeditor,
} = require("../middleware/authMiddleware");

router.post("/login", userController.login);
router.post("/create", userController.addNewUser);

module.exports = router;
