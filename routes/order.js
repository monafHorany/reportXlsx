const express = require("express");
const router = express.Router();
const orderController = require("../controller/orders");

router.post(
  "/fetchAllOrderFromWoocommerce",
  orderController.fetchAllSalesOrderFromWoocommerce
);
router.post(
  "/fetchAllRefundOrderFromWoocommerce",
  orderController.fetchAllRefundOrderFromWoocommerce
);
router.get("/saleDownload", orderController.downloadSaleFile);
router.get("/refundDownload", orderController.downloadRefundFile);
router.post("/statusHandler", orderController.orderStatusHandler);
router.get("/fetchAllNewOrder", orderController.fetchAllNewOrder);
router.get("/woo_order", orderController.woo_order);
// router.get("/saleReport", orderController.saleReport);
// router.get("/refundReport", orderController.refundReport);

module.exports = router;
