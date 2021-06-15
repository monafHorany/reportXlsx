const { Sequelize } = require("sequelize");
const dotenv = require("dotenv");
dotenv.config();

const sequelize = new Sequelize(
  process.env.DB_HOST,
  process.env.DB_USER,
  process.env.DB_PASS,
  {
    host: "127.0.0.1",
    dialect: "mysql",
    // logging: false,
  }
);

module.exports = sequelize;
