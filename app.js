const dotenv = require("dotenv");
dotenv.config();
const express = require("express");
const app = express();
app.use(express.json());
const path = require("path");
const cors = require("cors");
app.use(cors());
const sequelize = require("./utils/databaseConnection");

const ConfirmedOrder = require("./models/confirmed_order");

const CancelledOrder = require("./models/cancelled_order");

app.use("/order", require("./routes/order"));
app.use("/user", require("./routes/user-route"));

sequelize.sync().then(app.listen(8080));
