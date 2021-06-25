const express = require("express");
const router = express.Router();
const orderController = require("../controller/orders");

const {
  protect,
  accountant,
  ordermanager,
} = require("../middleware/authMiddleware");

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
router.post(
  "/statusHandler",
  protect,
  ordermanager,
  orderController.orderStatusHandler
);
router.get(
  "/fetchAllNewOrder",
  protect,
  accountant,
  orderController.fetchAllNewOrder
);
router.get("/woo_order", protect, accountant, orderController.woo_order);
// router.get("/saleReport", orderController.saleReport);
// router.get("/refundReport", orderController.refundReport);

module.exports = router;
