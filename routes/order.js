const express = require("express");
const router = express.Router();
const orderController = require("../controller/orders");

router.post(
  "/fetchAllOrderFromWoocommerce",
  orderController.fetchAllOrderFromWoocommerce
);
router.get("/download", orderController.downloadFiles);
router.post("/statusHandler", orderController.orderStatusHandler);
router.get("/fetchAllNewOrder", orderController.fetchAllNewOrder);
router.get("/woo_order", orderController.woo_order);
router.get("/saleReport", orderController.saleReport);
router.get("/refundReport", orderController.refundReport);

module.exports = router;
