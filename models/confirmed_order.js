const { DataTypes } = require("sequelize");
const sequelize = require("../utils/databaseConnection");

const ConfirmedOrder = sequelize.define("confirmed_Order", {
  id: {
    type: DataTypes.INTEGER,
    allowNull: false,
    autoIncrement: true,
    primaryKey: true,
  },
  woo_order_id: {
    type: DataTypes.INTEGER,
    allowNull: false,
  },
  orjeen_sku: {
    type: DataTypes.STRING,
    allowNull: false,
  },
  accountant_sku: {
    type: DataTypes.STRING,
    allowNull: true,
  },
  order_item_name: {
    type: DataTypes.STRING,
    allowNull: false,
  },
  quantity: {
    type: DataTypes.INTEGER,
    allowNull: false,
  },
  order_status: {
    type: DataTypes.STRING,
    allowNull: false,
  },
  shipping_method: {
    type: DataTypes.STRING,
    allowNull: false,
  },
  price: {
    type: DataTypes.DOUBLE,
    allowNull: false,
  },
  tax: {
    type: DataTypes.INTEGER,
    allowNull: false,
  },
  total: {
    type: DataTypes.DOUBLE,
    allowNull: false,
  },
  payment_method: {
    type: DataTypes.STRING,
    allowNull: false,
  },
  order_created_date: {
    type: DataTypes.DATE,
    allowNull: false,
  },
  order_modified_date: {
    type: DataTypes.DATE,
    allowNull: false,
  },
});

module.exports = ConfirmedOrder;
