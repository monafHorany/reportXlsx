const WooCommerceRestApi = require("@woocommerce/woocommerce-rest-api").default;
const asyncHandler = require("express-async-handler");
const moment = require("moment");
const excelToJson = require("convert-excel-to-json");
const csv = require("csvtojson");
// const { parse } = require("json2csv");
const AdmZip = require("adm-zip");
const zip = require("express-zip");
const json2xls = require("json2xls");
const fs = require("fs");
const ConfirmedOrder = require("../models/confirmed_order");

const CancelledOrder = require("../models/cancelled_order");
const ShippedOrder = require("../models/shipped_order");
const path = require("path");
var pdf = require("html-pdf");
const sequelize = require("../utils/databaseConnection");
const WooCommerce = new WooCommerceRestApi({
  url: "https://www.orjeen.com/",
  consumerKey: "ck_00ad078dc4ab7e0c25c7f277ea6f177d8fa17599",
  consumerSecret: "cs_6a20eea55b2b6dc379c98c3216cc1e4b7d9d74db",
  version: "wc/v3",
});
const fetchAllNewOrder = asyncHandler(async (req, res, next) => {
  let allOrders;
  try {
    allOrders = await Order.findAll({
      include: OrderItem,
      order: [["woo_order_id", "DESC"]],
    });
    return res.status(200).json(allOrders);
  } catch (error) {
    throw new Error(error);
  }
});

const woo_order = asyncHandler(async (req, res, next) => {
  const { data } = await WooCommerce.get(`products/categories?per_page=100`);
  let skus = "";
  return res.status(200).json(data);
});

const saleReport = asyncHandler(async (req, res, next) => {
  let confirmedOrder;
  let source = [];
  try {
    confirmedOrder = await ConfirmedOrder.findAll();
    confirmedOrder.forEach((order) => {
      source.push({
        woo_order_id: order.woo_order_id,
        orjeen_sku: order.orjeen_sku,
        accountant_sku: order.accountant_sku,
        order_item_name: order.order_item_name,
        quantity: order.quantity,
        order_status: order.order_status,
        shipping_method: order.shipping_method,
        price: order.price,
        tax: order.tax,
        total: order.total,
        payment_method: order.payment_method,
        order_created_date: new Date(order.order_created_date).toLocaleString(),
        order_modified_date: new Date(
          order.order_modified_date
        ).toLocaleString(),
      });
    });
  } catch (error) {
    throw new Error(error);
  }
  try {
    var xls = json2xls(source);
    fs.writeFileSync("./saleReport.xlsx", xls, "binary");
  } catch (error) {
    throw new Error(error);
  }
  return res.status(200).json(source);
});
const refundReport = asyncHandler(async (req, res, next) => {
  let confirmedOrder;
  let source = [];
  try {
    confirmedOrder = await CancelledOrder.findAll();
    confirmedOrder.forEach((order) => {
      source.push({
        woo_order_id: order.woo_order_id,
        orjeen_sku: order.orjeen_sku,
        accountant_sku: order.accountant_sku,
        order_item_name: order.order_item_name,
        quantity: order.quantity,
        order_status: order.order_status,
        shipping_method: order.shipping_method,
        price: order.price,
        tax: order.tax,
        total: order.total,
        payment_method: order.payment_method,
        order_created_date: new Date(order.order_created_date).toLocaleString(),
        order_modified_date: new Date(
          order.order_modified_date
        ).toLocaleString(),
      });
    });
  } catch (error) {
    throw new Error(error);
  }
  try {
    var xls = json2xls(source);
    fs.writeFileSync("./refundReport.xlsx", xls, "binary");
  } catch (error) {
    throw new Error(error);
  }
  return res.status(200).json(confirmedOrder);
});

const fetchAllSalesOrderFromWoocommerce = asyncHandler(
  async (req, res, next) => {
    const { from, to } = req.body;

    console.log(
      new Date(new Date(from).toISOString().substr(0, 10)).toISOString()
    );
    console.log(
      new Date(new Date(to).toISOString().substr(0, 10)).toISOString()
    );
    const jsonArray = excelToJson({
      sourceFile: "sku.xlsx",
      header: {
        rows: 1,
      },
      columnToKey: {
        A: "ACCsku",
        B: "ORsku",
      },
    });
    await sequelize.query("SET FOREIGN_KEY_CHECKS = 0");
    await ConfirmedOrder.destroy({
      truncate: true,
    });
    let page_num = 1;
    while (true) {
      const { data } = await WooCommerce.get(
        `orders?page=${page_num}&per_page=100&after=${new Date(
          from
        ).toISOString()}&before=${new Date(to).toISOString()}`
      );
      let createdOrder;
      for (let j = 0; j < data.length; j++) {
        const woo_order = data[j];
        if (woo_order.status) {
          try {
            const result = await sequelize.transaction(async (t) => {
              // if (woo_order.refunds.length === 0) {
              for (k = 0; k < woo_order.line_items.length; k++) {
                const orderItem = woo_order.line_items[k];
                if (orderItem.meta_data[0]) {
                  if (orderItem.meta_data[0].key) {
                    if (orderItem.meta_data[0].key != "_bundled_by") {
                      for (let q = 0; q < orderItem.quantity; q++) {
                        createdOrder = await ConfirmedOrder.create(
                          {
                            woo_order_id: woo_order.id,
                            orjeen_sku: orderItem.sku ? orderItem.sku : "N/A",
                            order_owner_name:
                              woo_order.billing.first_name +
                              " " +
                              woo_order.billing.last_name,
                            accountant_sku:
                              jsonArray.Sheet1[
                                jsonArray.Sheet1.findIndex(
                                  (s) => s.ORsku == orderItem.sku
                                )
                              ] &&
                              jsonArray.Sheet1[
                                jsonArray.Sheet1.findIndex(
                                  (s) => s.ORsku == orderItem.sku
                                )
                              ].ACCsku,
                            order_item_name: orderItem.name,
                            quantity: 1,
                            order_status: woo_order.status,
                            shipping_method: "SMSA_EXPRESS",
                            shipping_total: woo_order.shipping_total,
                            price: +orderItem.price.toFixed(2),
                            tax: +((orderItem.price * 15) / 100).toFixed(2),
                            payment_method: woo_order.payment_method_title,
                            order_created_date: woo_order.date_created,
                            order_modified_date: woo_order.date_modified,
                            item_sku: orderItem.sku,
                            item_price: orderItem.price,
                            item_quantity: orderItem.quantity,
                            total: +(
                              orderItem.price +
                              (orderItem.price * 15) / 100
                            ).toFixed(2),
                          },
                          { transaction: t }
                        );
                      }
                    }
                  }
                } else {
                  for (k = 0; k < woo_order.line_items.length; k++) {
                    const orderItem = woo_order.line_items[k];
                    for (let q = 0; q < orderItem.quantity; q++) {
                      createdOrder = await ConfirmedOrder.create(
                        {
                          woo_order_id: woo_order.id,
                          orjeen_sku: orderItem.sku ? orderItem.sku : "N/A",
                          order_owner_name:
                            woo_order.billing.first_name +
                            " " +
                            woo_order.billing.last_name,
                          accountant_sku:
                            jsonArray.Sheet1[
                              jsonArray.Sheet1.findIndex(
                                (s) => s.ORsku == orderItem.sku
                              )
                            ] &&
                            jsonArray.Sheet1[
                              jsonArray.Sheet1.findIndex(
                                (s) => s.ORsku == orderItem.sku
                              )
                            ].ACCsku,
                          order_item_name: orderItem.name,
                          quantity: 1,
                          order_status: "processing",
                          shipping_method: "SMSA_EXPRESS",
                          shipping_total: woo_order.shipping_total,
                          price: +orderItem.price.toFixed(2),
                          tax: +((orderItem.price * 15) / 100).toFixed(2),
                          payment_method: woo_order.payment_method_title,
                          order_created_date: woo_order.date_created,
                          order_modified_date: woo_order.date_modified,
                          item_sku: orderItem.sku,
                          item_price: orderItem.price,
                          item_quantity: orderItem.quantity,
                          total: +(
                            orderItem.price +
                            (orderItem.price * 15) / 100
                          ).toFixed(2),
                        },
                        { transaction: t }
                      );
                    }
                  }
                }
              }
            });
          } catch (error) {
            throw new Error(error);
          }
        }
      }
      page_num++;
      if (data.length === 0) {
        let salesSource = [];

        try {
          const confirmedOrder = await ConfirmedOrder.findAll();
          confirmedOrder.forEach((order) => {
            salesSource.push({
              woo_order_id: order.woo_order_id,
              orjeen_sku: order.orjeen_sku,
              ordered_by: order.order_owner_name,
              accountant_sku: order.accountant_sku,
              order_item_name: order.order_item_name,
              quantity: order.quantity,
              order_status: order.order_status,
              shipping_method: order.shipping_method,
              shipping_TAX: order.shipping_total,
              price: order.price,
              tax: order.tax,
              total: order.total,
              payment_method: order.payment_method,
              order_created_date: new Date(
                order.order_created_date
              ).toLocaleDateString(),
              order_modified_date: new Date(
                order.order_modified_date
              ).toLocaleDateString(),
            });
          });
        } catch (error) {
          throw new Error(error);
        }
        try {
          var xls = json2xls(salesSource);
          fs.writeFileSync("./saleReport.xlsx", xls, "binary");
        } catch (error) {
          throw new Error(error);
        }
        return res.status(200).json("being downloaded");
      }
    }
  }
);

const fetchAllRefundOrderFromWoocommerce = asyncHandler(
  async (req, res, next) => {
    const { from, to } = req.body;
    // console.log(req.body);
    // return res.json(req.body);
    const jsonArray = excelToJson({
      sourceFile: "sku.xlsx",
      header: {
        rows: 1,
      },
      columnToKey: {
        A: "ACCsku",
        B: "ORsku",
      },
    });
    await sequelize.query("SET FOREIGN_KEY_CHECKS = 0");
    await CancelledOrder.destroy({
      truncate: true,
    });
    let page_num = 1;
    while (true) {
      const { data } = await WooCommerce.get(
        `orders?page=${page_num}&per_page=100`
      );
      for (let j = 0; j < data.length; j++) {
        const woo_order = data[j];
        if (
          (woo_order.status === "returned-to-store" ||
            woo_order.status === "damaged-return" ||
            woo_order.status === "refunded" ||
            woo_order.status === "cancelled") &&
          +new Date(woo_order.date_modified) >= +new Date(from) &&
          +new Date(woo_order.date_modified) <= +new Date(to)
        ) {
          try {
            const result = await sequelize.transaction(async (t) => {
              for (k = 0; k < woo_order.line_items.length; k++) {
                const orderItem = woo_order.line_items[k];
                if (orderItem.meta_data[0]) {
                  if (orderItem.meta_data[0].key) {
                    if (orderItem.meta_data[0].key != "_bundled_by") {
                      for (let q = 0; q < orderItem.quantity; q++) {
                        const cancelled_order = await CancelledOrder.create(
                          {
                            woo_order_id: woo_order.id,
                            orjeen_sku: orderItem.sku ? orderItem.sku : "N/A",
                            order_owner_name:
                              woo_order.billing.first_name +
                              " " +
                              woo_order.billing.last_name,
                            accountant_sku:
                              jsonArray.Sheet1[
                                jsonArray.Sheet1.findIndex(
                                  (s) => s.ORsku == orderItem.sku
                                )
                              ] &&
                              jsonArray.Sheet1[
                                jsonArray.Sheet1.findIndex(
                                  (s) => s.ORsku == orderItem.sku
                                )
                              ].ACCsku,
                            order_item_name: orderItem.name
                              ? orderItem.name
                              : "N/A",
                            quantity: 1,
                            order_status: woo_order.status,
                            shipping_method: "SMSA_EXPRESS",
                            shipping_total: woo_order.shipping_total,
                            price: +orderItem.price.toFixed(2),
                            tax: +((orderItem.price * 15) / 100).toFixed(2),
                            payment_method: woo_order.payment_method_title,
                            order_created_date: woo_order.date_created,
                            order_modified_date: woo_order.date_modified,
                            item_sku: orderItem.sku,
                            item_price: orderItem.price,
                            item_quantity: orderItem.quantity,
                            total: +(
                              orderItem.price +
                              (orderItem.price * 15) / 100
                            ).toFixed(2),
                          },
                          { transaction: t }
                        );
                      }
                    }
                  }
                } else {
                  for (let q = 0; q < orderItem.quantity; q++) {
                    createdOrder = await CancelledOrder.create(
                      {
                        woo_order_id: woo_order.id,
                        orjeen_sku: orderItem.sku ? orderItem.sku : "N/A",
                        order_owner_name:
                          woo_order.billing.first_name +
                          " " +
                          woo_order.billing.last_name,
                        accountant_sku:
                          jsonArray.Sheet1[
                            jsonArray.Sheet1.findIndex(
                              (s) => s.ORsku == orderItem.sku
                            )
                          ] &&
                          jsonArray.Sheet1[
                            jsonArray.Sheet1.findIndex(
                              (s) => s.ORsku == orderItem.sku
                            )
                          ].ACCsku,
                        order_item_name: orderItem.name
                          ? orderItem.name
                          : "N/A",
                        quantity: 1,
                        order_status: woo_order.status,
                        shipping_method: "SMSA_EXPRESS",
                        shipping_total: woo_order.shipping_total,
                        price: +orderItem.price.toFixed(2),
                        tax: +((orderItem.price * 15) / 100).toFixed(2),
                        payment_method: woo_order.payment_method_title,
                        order_created_date: woo_order.date_created,
                        order_modified_date: woo_order.date_modified,
                        item_sku: orderItem.sku,
                        item_price: orderItem.price,
                        item_quantity: orderItem.quantity,
                        total: +(
                          orderItem.price +
                          (orderItem.price * 15) / 100
                        ).toFixed(2),
                      },
                      { transaction: t }
                    );
                  }
                }
              }
            });
          } catch (err) {
            throw new Error(err);
          }
        } else if (
          woo_order.status === "processing" ||
          woo_order.status === "completed" ||
          woo_order.status === "pending" ||
          woo_order.status === "shipped"
        ) {
          try {
            const result = await sequelize.transaction(async (t) => {
              if (woo_order.refunds.length !== 0) {
                for (k = 0; k < woo_order.line_items.length; k++) {
                  const orderItem = woo_order.line_items[k];
                  console.log(woo_order.id);
                  for (
                    let q = 0;
                    q <
                    orderItem.quantity -
                      (orderItem.meta_data.length > 0
                        ? orderItem.meta_data[0].value
                        : 0);
                    q++
                  ) {
                    createdOrder = await CancelledOrder.create(
                      {
                        woo_order_id: woo_order.id,
                        orjeen_sku: orderItem.sku ? orderItem.sku : "N/A",
                        order_owner_name:
                          woo_order.billing.first_name +
                          " " +
                          woo_order.billing.last_name,
                        accountant_sku:
                          jsonArray.Sheet1[
                            jsonArray.Sheet1.findIndex(
                              (s) => s.ORsku == orderItem.sku
                            )
                          ] &&
                          jsonArray.Sheet1[
                            jsonArray.Sheet1.findIndex(
                              (s) => s.ORsku == orderItem.sku
                            )
                          ].ACCsku,
                        order_item_name: orderItem.name
                          ? orderItem.name
                          : "N/A",
                        quantity: 1,
                        order_status: "refunded",
                        shipping_method: "SMSA_EXPRESS",
                        shipping_total: woo_order.shipping_total,
                        price: +orderItem.price.toFixed(2),
                        tax: +((orderItem.price * 15) / 100).toFixed(2),
                        payment_method: woo_order.payment_method_title,
                        order_created_date: woo_order.date_created,
                        order_modified_date: woo_order.date_modified,
                        item_sku: orderItem.sku,
                        item_price: orderItem.price,
                        item_quantity: orderItem.quantity,
                        total: +(
                          orderItem.price +
                          (orderItem.price * 15) / 100
                        ).toFixed(2),
                      },
                      { transaction: t }
                    );
                  }
                }
              }
            });
          } catch (error) {
            throw new Error(error);
          }
        }
      }
      page_num++;
      if (data.length === 0) {
        let refundSource = [];
        try {
          const cancelled = await CancelledOrder.findAll();
          cancelled.forEach((order) => {
            refundSource.push({
              woo_order_id: order.woo_order_id,
              orjeen_sku: order.orjeen_sku,
              ordered_by: order.order_owner_name,
              accountant_sku: order.accountant_sku,
              order_item_name: order.order_item_name,
              quantity: order.quantity,
              order_status: order.order_status,
              shipping_method: order.shipping_method,
              shipping_TAX: order.shipping_total,
              price: order.price,
              tax: order.tax,
              total: order.total,
              payment_method: order.payment_method,
              order_created_date: new Date(
                order.order_created_date
              ).toLocaleDateString(),
              order_modified_date: new Date(
                order.order_modified_date
              ).toLocaleDateString(),
            });
          });
        } catch (error) {
          throw new Error(error);
        }
        try {
          var xls = json2xls(refundSource);
          fs.writeFileSync("./refundReport.xlsx", xls, "binary");
        } catch (error) {
          throw new Error(error);
        }
        return res.status(200).json("being downloaded");
      }
    }
  }
);

const fetchAllShippedOrderFromWoocommerce = asyncHandler(
  async (req, res, next) => {
    const { from, to } = req.body;
    // console.log(req.body);
    // return res.json(req.body);
    const jsonArray = excelToJson({
      sourceFile: "sku.xlsx",
      header: {
        rows: 1,
      },
      columnToKey: {
        A: "ACCsku",
        B: "ORsku",
      },
    });
    await sequelize.query("SET FOREIGN_KEY_CHECKS = 0");
    await ShippedOrder.destroy({
      truncate: true,
    });
    let page_num = 1;
    while (true) {
      const { data } = await WooCommerce.get(
        `orders?page=${page_num}&per_page=100`
      );
      for (let j = 0; j < data.length; j++) {
        const woo_order = data[j];
        if (
          woo_order.status === "shipped" &&
          +new Date(woo_order.date_modified) >= +new Date(from) &&
          +new Date(woo_order.date_modified) <= +new Date(to)
        ) {
          try {
            const result = await sequelize.transaction(async (t) => {
              for (k = 0; k < woo_order.line_items.length; k++) {
                const orderItem = woo_order.line_items[k];
                if (orderItem.meta_data[0]) {
                  if (orderItem.meta_data[0].key) {
                    if (orderItem.meta_data[0].key != "_bundled_by") {
                      for (let q = 0; q < orderItem.quantity; q++) {
                        const shipped_order = await ShippedOrder.create(
                          {
                            woo_order_id: woo_order.id,
                            orjeen_sku: orderItem.sku ? orderItem.sku : "N/A",
                            order_owner_name:
                              woo_order.billing.first_name +
                              " " +
                              woo_order.billing.last_name,
                            accountant_sku:
                              jsonArray.Sheet1[
                                jsonArray.Sheet1.findIndex(
                                  (s) => s.ORsku == orderItem.sku
                                )
                              ] &&
                              jsonArray.Sheet1[
                                jsonArray.Sheet1.findIndex(
                                  (s) => s.ORsku == orderItem.sku
                                )
                              ].ACCsku,
                            order_item_name: orderItem.name
                              ? orderItem.name
                              : "N/A",
                            quantity: 1,
                            order_status: woo_order.status,
                            shipping_method: "SMSA_EXPRESS",
                            shipping_total: woo_order.shipping_total,
                            price: +orderItem.price.toFixed(2),
                            tax: +((orderItem.price * 15) / 100).toFixed(2),
                            payment_method: woo_order.payment_method_title,
                            order_created_date: woo_order.date_created,
                            order_modified_date: woo_order.date_modified,
                            item_sku: orderItem.sku,
                            item_price: orderItem.price,
                            item_quantity: orderItem.quantity,
                            total: +(
                              orderItem.price +
                              (orderItem.price * 15) / 100
                            ).toFixed(2),
                          },
                          { transaction: t }
                        );
                      }
                    }
                  }
                } else {
                  for (k = 0; k < woo_order.line_items.length; k++) {
                    const orderItem = woo_order.line_items[k];
                    for (let q = 0; q < orderItem.quantity; q++) {
                      const shipped_order = await ShippedOrder.create(
                        {
                          woo_order_id: woo_order.id,
                          orjeen_sku: orderItem.sku ? orderItem.sku : "N/A",
                          order_owner_name:
                            woo_order.billing.first_name +
                            " " +
                            woo_order.billing.last_name,
                          accountant_sku:
                            jsonArray.Sheet1[
                              jsonArray.Sheet1.findIndex(
                                (s) => s.ORsku == orderItem.sku
                              )
                            ] &&
                            jsonArray.Sheet1[
                              jsonArray.Sheet1.findIndex(
                                (s) => s.ORsku == orderItem.sku
                              )
                            ].ACCsku,
                          order_item_name: orderItem.name
                            ? orderItem.name
                            : "N/A",
                          quantity: 1,
                          order_status: woo_order.status,
                          shipping_method: "SMSA_EXPRESS",
                          shipping_total: woo_order.shipping_total,
                          price: +orderItem.price.toFixed(2),
                          tax: +((orderItem.price * 15) / 100).toFixed(2),
                          payment_method: woo_order.payment_method_title,
                          order_created_date: woo_order.date_created,
                          order_modified_date: woo_order.date_modified,
                          item_sku: orderItem.sku,
                          item_price: orderItem.price,
                          item_quantity: orderItem.quantity,
                          total: +(
                            orderItem.price +
                            (orderItem.price * 15) / 100
                          ).toFixed(2),
                        },
                        { transaction: t }
                      );
                    }
                  }
                }
              }
            });
          } catch (err) {
            throw new Error(err);
          }
        }
      }
      page_num++;
      if (data.length === 0) {
        let shippedSource = [];
        try {
          const shipped = await ShippedOrder.findAll();
          shipped.forEach((order) => {
            shippedSource.push({
              woo_order_id: order.woo_order_id,
              orjeen_sku: order.orjeen_sku,
              ordered_by: order.order_owner_name,
              accountant_sku: order.accountant_sku,
              order_item_name: order.order_item_name,
              quantity: order.quantity,
              order_status: order.order_status,
              shipping_method: order.shipping_method,
              shipping_TAX: order.shipping_total,
              price: order.price,
              tax: order.tax,
              total: order.total,
              payment_method: order.payment_method,
              order_created_date: new Date(
                order.order_created_date
              ).toLocaleDateString(),
              order_modified_date: new Date(
                order.order_modified_date
              ).toLocaleDateString(),
            });
          });
        } catch (error) {
          throw new Error(error);
        }
        try {
          var xls = json2xls(shippedSource);
          fs.writeFileSync("./shippedReport.xlsx", xls, "binary");
        } catch (error) {
          throw new Error(error);
        }
        return res.status(200).json("being downloaded");
      }
    }
  }
);

const downloadSaleFile = asyncHandler(async (req, res, next) => {
  const invoiceName = "saleReport.xlsx";
  const invoicePath = path.resolve(invoiceName);
  fs.readFile(invoicePath, (err, data) => {
    if (err) {
      return console.log("err" + err);
    }
    res.setHeader("Content-Type", "application/xlsx");
    res.setHeader(
      "Content-Disposition",
      'inline; filename="' + invoiceName + '"'
    );
    return res.status(200).send(data);
  });
});
const downloadRefundFile = asyncHandler(async (req, res, next) => {
  const invoiceName = "refundReport.xlsx";
  const invoicePath = path.resolve(invoiceName);
  fs.readFile(invoicePath, (err, data) => {
    if (err) {
      return console.log("err" + err);
    }
    res.setHeader("Content-Type", "application/xlsx");
    res.setHeader(
      "Content-Disposition",
      'inline; filename="' + invoiceName + '"'
    );
    return res.status(200).send(data);
  });
});

const downloadShippedFile = asyncHandler(async (req, res, next) => {
  const invoiceName = "shippedReport.xlsx";
  const invoicePath = path.resolve(invoiceName);
  fs.readFile(invoicePath, (err, data) => {
    if (err) {
      return console.log("err" + err);
    }
    res.setHeader("Content-Type", "application/xlsx");
    res.setHeader(
      "Content-Disposition",
      'inline; filename="' + invoiceName + '"'
    );
    return res.status(200).send(data);
  });
});

const orderStatusHandler = asyncHandler(async (req, res, next) => {
  const { ids, status } = req.body;
  // return console.log(req.body);
  console.log(ids.split("\n"));
  // return res.status(201).json("done");
  const update = {
    status: status,
  };
  if (ids) {
    const convertedIds = ids.split("\n");
    console.log(convertedIds);
    // ids.split("\n").map((id, index) => {
    //   console.log(parseInt(id));
    //   if (!isNaN(parseInt(id))) {
    //     const { data } = await WooCommerce.put(
    //       `orders/${parseInt(id)}`,
    //       update
    //     );
    //   }
    //   console.log(index);
    //   if (index === ids.split("\n").length - 1) {
    //     return res.status(201).json("done");
    //   }
    // });
  }
});
const changePriceHandler = asyncHandler(async (req, res, next) => {
  return console.log(req.body);
  const { ids } = req.body;
  console.log(ids.split("\n"));
  // return res.status(201).json("done");
  const update = {
    status: "cancelled",
  };
  if (ids) {
    ids.split("\n").map(async (id, index) => {
      console.log(parseInt(id));
      if (!isNaN(parseInt(id))) {
        const { data } = await WooCommerce.put(
          `orders/${parseInt(id)}`,
          update
        );
      }
      console.log(index);
      if (index === ids.split("\n").length - 1) {
        return res.status(201).json("done");
      }
    });
  }
});

const printBulkInvoice = asyncHandler(async (req, res, next) => {
  // const { ids } = req.body;
  const response = [];
  // const ids = [62953, 62931, 62902];
  const ids = [52242];
  // if (ids) {
  //ids.split("\n")
  try {
    ids.map(async (id, index) => {
      // if (!isNaN(parseInt(id))) {
      // const { data } = await WooCommerce.get(`orders/${id}`);
      const { data } = await WooCommerce.get(`orders/${id}`);
      const options = {
        format: "letter",
        orientation: "landscape",
        // border: "2mm",
      };
      const document = `<!DOCTYPE html>
        <html>
          <head>
          
          <style>
          .flex-container {
            display: -webkit-box;
            display: -ms-flexbox;
            display: flex;
            background-color: DodgerBlue;
          }
          
          .flex-container > div {
            background-color: #f1f1f1;
            margin: 10px;
            padding: 20px;
            font-size: 30px;
          }
          </style>
          </head>
          <body>
          
          <div class="flex-container">
            <div>1</div>
            <div>2</div>
            <div>3</div>  
          </div>
          </body>
        </html>
        `;
      pdf
        .create(document, options)
        .toFile(`./${data.id}.pdf`, function (err, res) {
          const invoiceName = `${data.id}.pdf`;
          const invoicePath = path.resolve(invoiceName);
          fs.readFile(invoicePath, (err, data) => {
            if (err) {
              return console.log("err" + err);
            }
          });
        });

      response.push(data);
      // }
      // console.log(index);
      if (index === ids.length - 1) return res.status(200).send(response);
      // if (index === ids.split("\n").length - 1) {
      // }
    });
    // }
  } catch (error) {
    console.error(error);
  }
});

exports.fetchAllSalesOrderFromWoocommerce = fetchAllSalesOrderFromWoocommerce;
exports.fetchAllRefundOrderFromWoocommerce = fetchAllRefundOrderFromWoocommerce;
exports.fetchAllShippedOrderFromWoocommerce =
  fetchAllShippedOrderFromWoocommerce;
exports.fetchAllNewOrder = fetchAllNewOrder;
exports.woo_order = woo_order;
// exports.saleReport = saleReport;
// exports.refundReport = refundReport;
exports.downloadSaleFile = downloadSaleFile;
exports.downloadRefundFile = downloadRefundFile;
exports.downloadShippedFile = downloadShippedFile;
exports.orderStatusHandler = orderStatusHandler;
exports.changePriceHandler = changePriceHandler;
exports.printBulkInvoice = printBulkInvoice;
