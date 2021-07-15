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
const path = require("path");
const pdf = require("pdf-creator-node");
const sequelize = require("../utils/databaseConnection");
const WooCommerce = new WooCommerceRestApi({
  url: "https://www.orjeen.com/",
  consumerKey: "ck_6a4943efca5beb973900d31d6a5e4f397c8116ba",
  consumerSecret: "cs_85a3dcb8a42ce7b2c4714bb1d6027b3196c8bc8e",
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
  const { data } = await WooCommerce.get(`orders/55166`);

  let skus = "";
  // for (let index = 0; index < data.line_items.length; index++) {
  //   const item = data.line_items[index];
  //   console.log(item);
  //   console.log(
  //     item.meta_data[0].key === "_bundled_items" &&
  //       item.bundled_items.length > 0
  //   );
  //   if (
  //     item.meta_data[0].key === "_bundled_items" &&
  //     item.bundled_items.length > 0
  //   ) {
  //     for (let index2 = 0; index2 < item.bundled_items.length; index2++) {
  //       const bundled_item_id = item.bundled_items[index2];
  //       console.log(bundled_item_id);
  //       for (let index3 = 0; index3 < data.line_items.length; index3++) {
  //         for (let index4 = 0; index4 < item.quantity; index4++) {
  //           const i = data.line_items[index3];
  //           if (i.id === bundled_item_id) {
  //             skus += i.sku + ",";
  //           }
  //         }
  //       }
  //     }
  //   }
  // }
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
      sourceFile: "ACC SKU.xlsx",
      header: {
        rows: 1,
      },
      columnToKey: {
        A: "ACCsku",
        B: "ORsku",
      },
    });
    // console.log(jsonArray.Sheet1[211]);
    // let aksku = "";
    // console.log(
    //   ["RJ4000401", "RJ9000553", "RJ9000553"]
    //     .map((sku) => {
    //       return jsonArray.Sheet1[
    //         jsonArray.Sheet1.findIndex((s) => s.ORsku == sku)
    //       ];
    //     })
    //     .forEach((e) => (aksku += e.ACsku + ","))
    // );
    // console.log("aksku", aksku);
    // return res.json(jsonArray.Sheet1.findIndex((s) => s.ORsku == "RJ9000553"));

    // var threeMonthsAgo = moment().subtract(1, "weeks");
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
        if (
          woo_order.status === "processing" ||
          woo_order.status === "confirmed" ||
          woo_order.status === "completed"
        ) {
          try {
            const result = await sequelize.transaction(async (t) => {
              if (woo_order.refunds.length === 0) {
                for (k = 0; k < woo_order.line_items.length; k++) {
                  const orderItem = woo_order.line_items[k];
                  if (orderItem.meta_data[0]) {
                    if (orderItem.meta_data[0].key) {
                      if (orderItem.meta_data[0].key != "_bundled_by") {
                        for (let q = 0; q < orderItem.quantity; q++) {
                          createdOrder = await ConfirmedOrder.create(
                            {
                              woo_order_id: woo_order.id,
                              orjeen_sku: orderItem.sku,
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
                              shipping_method:
                                woo_order.shipping_lines[0].method_id,
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
                }
              } else {
                for (k = 0; k < woo_order.line_items.length; k++) {
                  const orderItem = woo_order.line_items[k];
                  for (let q = 0; q < orderItem.quantity; q++) {
                    createdOrder = await ConfirmedOrder.create(
                      {
                        woo_order_id: woo_order.id,
                        orjeen_sku: orderItem.sku,
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
                        shipping_method: woo_order.shipping_lines[0].method_id,
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
        let salesSource = [];

        try {
          const confirmedOrder = await ConfirmedOrder.findAll();
          confirmedOrder.forEach((order) => {
            salesSource.push({
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
      sourceFile: "ACC SKU.xlsx",
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
        `orders?page=${page_num}&per_page=100&status=cancelled,refunded,damaged-return,returned-to-store`
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
                            shipping_method:
                              woo_order.shipping_lines[0].method_id,
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
                        shipping_method: woo_order.shipping_lines[0].method_id,
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
          woo_order.status === "completed"
        ) {
          try {
            const result = await sequelize.transaction(async (t) => {
              // if (woo_order.refunds.length === 0) {
              //   for (k = 0; k < woo_order.line_items.length; k++) {
              //     const orderItem = woo_order.line_items[k];
              //     if (orderItem.meta_data[0]) {
              //       if (orderItem.meta_data[0].key) {
              //         if (orderItem.meta_data[0].key != "_bundled_by") {
              //           for (let q = 0; q < orderItem.quantity; q++) {
              //             createdOrder = await ConfirmedOrder.create(
              //               {
              //                 woo_order_id: woo_order.id,
              //                 orjeen_sku: orderItem.sku,
              //                 accountant_sku:
              //                   jsonArray.Sheet1[
              //                     jsonArray.Sheet1.findIndex(
              //                       (s) => s.ORsku == orderItem.sku
              //                     )
              //                   ] &&
              //                   jsonArray.Sheet1[
              //                     jsonArray.Sheet1.findIndex(
              //                       (s) => s.ORsku == orderItem.sku
              //                     )
              //                   ].ACCsku,
              //                 order_item_name: orderItem.name,
              //                 quantity: 1,
              //                 order_status: "processing",
              //                 shipping_method:
              //                   woo_order.shipping_lines[0].method_id,
              //                 price: +orderItem.price.toFixed(2),
              //                 tax: +((orderItem.price * 15) / 100).toFixed(2),
              //                 payment_method: woo_order.payment_method_title,
              //                 order_created_date: woo_order.date_created,
              //                 order_modified_date: woo_order.date_modified,
              //                 item_sku: orderItem.sku,
              //                 item_price: orderItem.price,
              //                 item_quantity: orderItem.quantity,
              //                 total: +(
              //                   orderItem.price +
              //                   (orderItem.price * 15) / 100
              //                 ).toFixed(2),
              //               },
              //               { transaction: t }
              //             );
              //           }
              //         }
              //       }
              //     }
              //   }
              // } else
              if (woo_order.refunds.length !== 0) {
                for (k = 0; k < woo_order.line_items.length; k++) {
                  const orderItem = woo_order.line_items[k];
                  for (
                    let q = 0;
                    q < orderItem.quantity - orderItem.meta_data[0].value;
                    q++
                  ) {
                    createdOrder = await CancelledOrder.create(
                      {
                        woo_order_id: woo_order.id,
                        orjeen_sku: orderItem.sku ? orderItem.sku : "N/A",
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
                        shipping_method: woo_order.shipping_lines[0].method_id,
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
              accountant_sku: order.accountant_sku,
              order_item_name: order.order_item_name,
              quantity: order.quantity,
              order_status: order.order_status,
              shipping_method: order.shipping_method,
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

const orderStatusHandler = asyncHandler(async (req, res, next) => {
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

exports.fetchAllSalesOrderFromWoocommerce = fetchAllSalesOrderFromWoocommerce;
exports.fetchAllRefundOrderFromWoocommerce = fetchAllRefundOrderFromWoocommerce;
exports.fetchAllNewOrder = fetchAllNewOrder;
exports.woo_order = woo_order;
// exports.saleReport = saleReport;
// exports.refundReport = refundReport;
exports.downloadSaleFile = downloadSaleFile;
exports.downloadRefundFile = downloadRefundFile;
exports.orderStatusHandler = orderStatusHandler;
exports.changePriceHandler = changePriceHandler;
