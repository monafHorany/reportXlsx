const { Sequelize, DataTypes } = require("sequelize");
const sequelize = require("../utils/databaseConnection");

const User = sequelize.define("user", {
  id: {
    type: DataTypes.INTEGER,
    autoIncrement: true,
    allowNull: false,
    primaryKey: true,
  },
  name: {
    type: DataTypes.STRING,
    allowNull: false,
  },
  password: {
    type: DataTypes.STRING,
    allowNull: false,
  },
  role: {
    type: DataTypes.ENUM,
    allowNull: false,
    values: ["accountant", "ordermanager", "productmanager"],
    validate: {
      isIn: {
        args: [["accountant", "ordermanager", "productmanager"]],
        msg: "Must be accountant or order manager or product manager",
      },
    },
  },
});

module.exports = User;
