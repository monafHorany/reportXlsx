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
  const { data } = await WooCommerce.get(`orders/52242`);

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
                            shipping_method: "SMSA_EXPRESS",
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
                          shipping_method: "SMSA_EXPRESS",
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
              // } else {
              //   for (k = 0; k < woo_order.line_items.length; k++) {
              //     const orderItem = woo_order.line_items[k];
              //     for (let q = 0; q < orderItem.quantity; q++) {
              //       createdOrder = await ConfirmedOrder.create(
              //         {
              //           woo_order_id: woo_order.id,
              //           orjeen_sku: orderItem.sku,
              //           accountant_sku:
              //             jsonArray.Sheet1[
              //               jsonArray.Sheet1.findIndex(
              //                 (s) => s.ORsku == orderItem.sku
              //               )
              //             ] &&
              //             jsonArray.Sheet1[
              //               jsonArray.Sheet1.findIndex(
              //                 (s) => s.ORsku == orderItem.sku
              //               )
              //             ].ACCsku,
              //           order_item_name: orderItem.name,
              //           quantity: 1,
              //           order_status: "processing",
              //           shipping_method: "SMSA_EXPRESS",
              //           price: +orderItem.price.toFixed(2),
              //           tax: +((orderItem.price * 15) / 100).toFixed(2),
              //           payment_method: woo_order.payment_method_title,
              //           order_created_date: woo_order.date_created,
              //           order_modified_date: woo_order.date_modified,
              //           item_sku: orderItem.sku,
              //           item_price: orderItem.price,
              //           item_quantity: orderItem.quantity,
              //           total: +(
              //             orderItem.price +
              //             (orderItem.price * 15) / 100
              //           ).toFixed(2),
              //         },
              //         { transaction: t }
              //       );
              //     }
              //   }
              // }
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
                            shipping_method: "SMSA_EXPRESS",
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
                        shipping_method: "SMSA_EXPRESS",
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
                        shipping_method: "SMSA_EXPRESS",
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
  const { ids, status } = req.body;
  // return console.log(req.body);
  console.log(ids.split("\n"));
  // return res.status(201).json("done");
  const update = {
    status: status,
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
        format: "A4",
        orientation: "landscape",
        // border: "2mm",
      };
      const document = {
        html: `
        
        <!DOCTYPE html PUBLIC "-//W3C//DTD html 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
<TITLE>pdf-html</TITLE>
<META name="generator" content="BCL easyConverter SDK 5.0.252">
<STYLE type="text/css">

body {display: flex; justify-content: center;}

#page_1 {position:relative; overflow: hidden;margin: 56px 0px 272px 0px;padding: 0px;border: none;width: 704px;height: 728px;}
#page_1 #id1_1 {border:none;margin: 10px 0px 0px 0px;padding: 0px;border:none;width: 688px;overflow: hidden;}
#page_1 #id1_2 {border:none;margin: 29px 0px 0px 113px;padding: 0px;border:none;width: 553px;overflow: hidden;}
#page_1 #id1_2 #id1_2_1 {float:left;border:none;margin: 2px 0px 0px 0px;padding: 0px;border:none;width: 311px;overflow: hidden;}
#page_1 #id1_2 #id1_2_2 {float:left;border:none;margin: 0px 0px 0px 47px;padding: 0px;border:none;width: 195px;overflow: hidden;}
#page_1 #id1_3 {border:none;margin: 57px 0px 0px 113px;padding: 0px;border:none;width: 591px;overflow: hidden;}
// #page_1 #id1_4 {border:none;margin: 66px 0px 0px 113px;padding: 0px;border:none;width: 511px;overflow: hidden;}
// #page_1 #id1_4 #id1_4_1 {float:left;border:none;margin: 0px 0px 0px 0px;padding: 0px;border:none;width: 284px;overflow: hidden;}
#page_1 #id1_4 #id1_4_2 {float:left;border:none;margin: 9px 0px 0px 111px;padding: 0px;border:none;width: 72px;overflow: hidden;}
#page_1 #id1_4 #id1_4_3 {float:left;border:none;margin: 3px 0px 0px 0px;padding: 0px;border:none;width: 44px;overflow: hidden;}

#page_1 #p1dimg1 {position:absolute;top:0px;left:113px;z-index:-1;width:553px;height:728px;}
#page_1 #p1dimg1 #p1img1 {width:553px;height:728px;}




.dclr {clear:both;float:none;height:1px;margin:0px;padding:0px;overflow:hidden;}

.ft0{font: bold 16px 'Arial';line-height: 19px;}
.ft1{font: 17px 'Arial';line-height: 19px;}
.ft2{font: bold 15px 'Arial';line-height: 18px;}
.ft3{font: 16px 'Arial';line-height: 18px;}
.ft4{font: 1px 'Arial';line-height: 14px;}
.ft5{font: 1px 'Arial';line-height: 1px;}
.ft6{font: bold 22px 'Arial';line-height: 26px;}
.ft7{font: bold 9px 'Arial';line-height: 11px;}
.ft8{font: bold 8px 'Arial';line-height: 10px;}
.ft9{font: 1px 'Arial';line-height: 6px;}
.ft10{font: 1px 'Arial';line-height: 7px;}
.ft11{font: 9px 'Arial';line-height: 12px;}
.ft12{font: 8px 'Arial';line-height: 10px;}
.ft13{font: bold 9px 'Arial';margin-left: 81px;line-height: 14px;}
.ft14{font: bold 9px 'Arial';line-height: 14px;}
.ft15{font: 9px 'Arial';line-height: 13px;}

.p0{text-align: center;padding-right: 29px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p1{text-align: center;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p2{text-align: center;padding-right: 33px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p3{text-align: left;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p4{text-align: left;padding-left: 23px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p5{text-align: right;padding-right: 9px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p6{text-align: left;padding-left: 10px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p7{text-align: right;margin-top: 0px;margin-bottom: 0px;}
.p8{text-align: left;padding-left: 58px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p9{text-align: left;padding-left: 43px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p10{text-align: center;padding-right: 13px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p11{text-align: center;padding-right: 15px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p12{text-align: center;padding-right: 16px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p13{text-align: center;padding-right: 18px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p14{text-align: center;padding-right: 12px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p15{text-align: right;padding-right: 55px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p16{text-align: center;padding-right: 61px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p17{text-align: center;padding-right: 1px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p18{text-align: left;padding-left: 25px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p19{text-align: center;padding-left: 2px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p20{text-align: center;padding-right: 60px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p21{text-align: left;padding-left: 20px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p22{text-align: left;padding-left: 14px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p23{text-align: left;padding-left: 8px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p24{text-align: left;padding-left: 13px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p25{text-align: left;padding-left: 11px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p26{text-align: left;padding-left: 5px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p27{text-align: left;padding-left: 47px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p28{text-align: left;padding-left: 19px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p29{text-align: left;padding-left: 6px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p30{text-align: left;padding-left: 9px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p31{text-align: right;padding-right: 5px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p32{text-align: right;padding-right: 11px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p33{text-align: center;padding-left: 1px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p34{text-align: center;padding-right: 24px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p35{text-align: center;padding-right: 22px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p36{text-align: center;padding-right: 25px;margin-top: 0px;margin-bottom: 0px;white-space: nowrap;}
.p37{text-align: left;padding-left: 142px;padding-right: 55px;margin-top: 11px;margin-bottom: 0px;text-indent: -119px;}
.p38{text-align: left;margin-top: 0px;margin-bottom: 0px;}

.td0{padding: 0px;margin: 0px;width: 196px;vertical-align: bottom;}
.td1{padding: 0px;margin: 0px;width: 189px;vertical-align: bottom;}
.td2{padding: 0px;margin: 0px;width: 72px;vertical-align: bottom;}
.td3{padding: 0px;margin: 0px;width: 124px;vertical-align: bottom;}
.td4{border-left: #000000 1px solid;border-top: #000000 1px solid;padding: 0px;margin: 0px;width: 164px;vertical-align: bottom;}
.td5{border-right: #000000 1px solid;border-top: #000000 1px solid;padding: 0px;margin: 0px;width: 146px;vertical-align: bottom;}
.td6{border-left: #000000 1px solid;padding: 0px;margin: 0px;width: 164px;vertical-align: bottom;}
.td7{border-right: #000000 1px solid;padding: 0px;margin: 0px;width: 146px;vertical-align: bottom;}
.td8{border-left: #000000 1px solid;border-bottom: #bbbbbb 1px solid;padding: 0px;margin: 0px;width: 164px;vertical-align: bottom;}
.td9{border-right: #000000 1px solid;border-bottom: #bbbbbb 1px solid;padding: 0px;margin: 0px;width: 146px;vertical-align: bottom;}
.td10{border-left: #000000 1px solid;border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 164px;vertical-align: bottom;}
.td11{border-right: #000000 1px solid;border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 146px;vertical-align: bottom;}
.td12{padding: 0px;margin: 0px;width: 0px;vertical-align: bottom;}
.td13{padding: 0px;margin: 0px;width: 130px;vertical-align: bottom;}
.td14{padding: 0px;margin: 0px;width: 65px;vertical-align: bottom;}
.td15{border-left: #000000 1px solid;border-right: #000000 1px solid;border-top: #000000 1px solid;padding: 0px;margin: 0px;width: 73px;vertical-align: bottom;}
.td16{border-right: #000000 1px solid;border-top: #000000 1px solid;padding: 0px;margin: 0px;width: 58px;vertical-align: bottom;}
.td17{border-right: #000000 1px solid;border-top: #000000 1px solid;padding: 0px;margin: 0px;width: 53px;vertical-align: bottom;}
.td18{border-right: #000000 1px solid;border-top: #000000 1px solid;padding: 0px;margin: 0px;width: 41px;vertical-align: bottom;}
.td19{border-right: #000000 1px solid;border-top: #000000 1px solid;padding: 0px;margin: 0px;width: 54px;vertical-align: bottom;}
.td20{border-right: #000000 1px solid;border-top: #000000 1px solid;padding: 0px;margin: 0px;width: 36px;vertical-align: bottom;}
.td21{border-right: #000000 1px solid;border-top: #000000 1px solid;padding: 0px;margin: 0px;width: 27px;vertical-align: bottom;}
.td22{border-right: #000000 1px solid;border-top: #000000 1px solid;padding: 0px;margin: 0px;width: 176px;vertical-align: bottom;}
.td23{border-right: #000000 1px solid;border-top: #000000 1px solid;padding: 0px;margin: 0px;width: 22px;vertical-align: bottom;background: #cccccc;}
.td24{border-left: #000000 1px solid;border-right: #000000 1px solid;padding: 0px;margin: 0px;width: 73px;vertical-align: bottom;}
.td25{border-right: #000000 1px solid;padding: 0px;margin: 0px;width: 58px;vertical-align: bottom;}
.td26{border-right: #000000 1px solid;padding: 0px;margin: 0px;width: 53px;vertical-align: bottom;}
.td27{border-right: #000000 1px solid;padding: 0px;margin: 0px;width: 41px;vertical-align: bottom;}
.td28{border-right: #000000 1px solid;padding: 0px;margin: 0px;width: 54px;vertical-align: bottom;}
.td29{border-right: #000000 1px solid;padding: 0px;margin: 0px;width: 36px;vertical-align: bottom;}
.td30{border-right: #000000 1px solid;padding: 0px;margin: 0px;width: 27px;vertical-align: bottom;}
.td31{border-right: #000000 1px solid;padding: 0px;margin: 0px;width: 176px;vertical-align: bottom;}
.td32{border-right: #000000 1px solid;padding: 0px;margin: 0px;width: 22px;vertical-align: bottom;background: #cccccc;}
.td33{border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 73px;vertical-align: bottom;}
.td34{border-right: #000000 1px solid;border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 58px;vertical-align: bottom;}
.td35{border-right: #000000 1px solid;border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 53px;vertical-align: bottom;}
.td36{border-right: #000000 1px solid;border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 41px;vertical-align: bottom;}
.td37{border-right: #000000 1px solid;border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 54px;vertical-align: bottom;}
.td38{border-right: #000000 1px solid;border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 36px;vertical-align: bottom;}
.td39{border-right: #000000 1px solid;border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 27px;vertical-align: bottom;}
.td40{border-right: #000000 1px solid;border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 176px;vertical-align: bottom;}
.td41{border-right: #000000 1px solid;border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 22px;vertical-align: bottom;background: #cccccc;}
.td42{border-right: #000000 1px solid;padding: 0px;margin: 0px;width: 22px;vertical-align: bottom;}
.td43{border-right: #000000 1px solid;border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 22px;vertical-align: bottom;}
.td44{padding: 0px;margin: 0px;width: 99px;vertical-align: bottom;}
.td45{padding: 0px;margin: 0px;width: 185px;vertical-align: bottom;}
.td46{border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 99px;vertical-align: bottom;}
.td47{border-bottom: #000000 1px solid;padding: 0px;margin: 0px;width: 185px;vertical-align: bottom;}

.tr0{height: 32px;}
.tr1{height: 24px;}
.tr2{height: 38px;}
.tr3{height: 14px;}
.tr4{height: 41px;}
.tr5{height: 60px;}
.tr6{height: 19px;}
.tr7{height: 22px;}
.tr8{height: 18px;}
.tr9{height: 6px;}
.tr10{height: 7px;}
.tr11{height: 20px;}
.tr12{height: 26px;}
.tr13{height: 13px;}
.tr14{height: 17px;}
.tr15{height: 16px;}
.tr16{height: 23px;}

.t0{width: 385px;margin-left: 298px;font: 16px 'Arial';}
.t1{width: 311px;font: bold 9px 'Arial';}
.t2{width: 195px;font: 9px 'Arial';}
.t3{width: 591px;font: 9px 'Arial';}
.t4{width: 284px;font: 9px 'Arial';}

</STYLE>
</HEAD>

<body>
<div id="page_1">
<div id="p1dimg1">
<img src="data:image/png;base64,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AGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQ1tjpHYC7UFmWO70Lt5dGw68a/oVGWZaNRuPu/l/lrv8Bua34721NO/6x7PgOUGs0Gs65AZCmPQCkaQ8AadoDQJr2AJCmPQCkaQ8AadoDQJr2AJCmPQCkaQ8AadoDQJr2AJCmPQCkaQ8AadoDQJr2AJCmPQCkaQ8AadoDQJr2AJCmPQCkaQ8AadoDQJr2AJCmPQCkaQ8AadoDQJr2AJCmPQCkaQ8AadoDQJr2AJCmPQCkaQ8Aabt3egeAu1+j0aj/ueO7we1Ae4BbrizLnd4Fbi/OuQGQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0Ca9gCQpj0ApGkPAGnaA0BaY6d3AO5CZVnu9C7AbW13WZaNRuPu/l/lrv8Bua00Gv5KB5twzg2ANO0BIE17AEjTHgDStAeANO0BIE17AEjTHgDStAeANO0BIE17AEjTHgDStAeANO0BIE17AEjTHgDStAeANO0BIE17AEjTHgDStAeANO0BIE17AEjTHgDStAeANO0BIE17AEjTHgDStAeANO0BIE17AEjTHgDStAeANO0BIE17AEhr7PQOwN2mLMud3gW43f0/5Iia9SFS1QIAAAAASUVORK5CYII=" id="p1img1"></div>


<div class="dclr"></div>
<div id="id1_1">
<TABLE cellpadding=0 cellspacing=0 class="t0">
<TR>
	<TD colspan=2 class="tr0 td0"><P class="p0 ft0">فاتورة مبيعات</P></TD>
	<TD class="tr0 td1"><P class="p1 ft1">شركة أول إتصال للتجارة والتسویق</P></TD>
</TR>
<TR>
	<TD colspan=2 class="tr1 td0"><P class="p2 ft2">TAX INVOICE</P></TD>
	<TD rowspan=2 class="tr2 td1"><P class="p1 ft3">(شركة شخص واحد)</P></TD>
</TR>
<TR>
	<TD class="tr3 td2"><P class="p3 ft4">&nbsp;</P></TD>
	<TD class="tr3 td3"><P class="p3 ft4">&nbsp;</P></TD>
</TR>
<TR>
	<TD class="tr4 td2"><P class="p3 ft5">&nbsp;</P></TD>
	<TD rowspan=2 class="tr5 td3"><P class="p4 ft0">:رقم الفاتورة</P></TD>
	<TD class="tr4 td1"><P class="p1 ft3">المملكة العربیة السعودیة</P></TD>
</TR>
<TR>
	<TD rowspan=2 class="tr4 td2"><P class="p5 ft6">${data.id}</P></TD>
	<TD class="tr6 td1"><P class="p1 ft3">الریاض -الملقا– 9152</P></TD>
</TR>
<TR>
	<TD class="tr7 td3"><P class="p6 ft0">Invoice No.</P></TD>
	<TD class="tr7 td1"><P class="p3 ft5">&nbsp;</P></TD>
</TR>
</TABLE>
<P class="p7 ft1">VAT No: 31053701050003</P>
</div>
<div id="id1_2">
<div id="id1_2_1">
<TABLE cellpadding=0 cellspacing=0 class="t1">
<TR>
	<TD class="tr8 td4"><P class="p8 ft7">اسم العمیل</P></TD>
	<TD class="tr8 td5"><P class="p9 ft7">${
    data.billing.first_name + " " + data.billing.last_name
  }</P></TD>
</TR>
<TR>
	<TD class="tr8 td6"><P class="p10 ft8">Customer Name</P></TD>
	<TD class="tr8 td7"><P class="p3 ft5">&nbsp;</P></TD>
</TR>
<TR>
	<TD class="tr9 td8"><P class="p3 ft9">&nbsp;</P></TD>
	<TD class="tr9 td9"><P class="p3 ft9">&nbsp;</P></TD>
</TR>
<TR>
	<TD class="tr8 td6"><P class="p11 ft7">عنوان العمیل</P></TD>
	<TD class="tr8 td7"><P class="p12 ft7">${
    data.shipping.address_1 +
    " ," +
    data.shipping.city +
    " ," +
    data.shipping.state +
    " ," +
    data.shipping.address_2
  }</P></TD>
</TR>
<TR>
	<TD class="tr8 td6"><P class="p10 ft7">Customer Address</P></TD>
	<TD class="tr8 td7"><P class="p3 ft5">&nbsp;</P></TD>
</TR>
<TR>
	<TD class="tr9 td8"><P class="p3 ft9">&nbsp;</P></TD>
	<TD class="tr9 td9"><P class="p3 ft9">&nbsp;</P></TD>
</TR>
<TR>
	<TD class="tr8 td6"><P class="p11 ft7">Email البرید الالكتروني</P></TD>
	<TD class="tr8 td7"><P class="p13 ft7">${data.billing.email}</P></TD>
</TR>
<TR>
	<TD class="tr8 td6"><P class="p14 ft8">Phone Number رقم الھاتف</P></TD>
	<TD class="tr8 td7"><P class="p15 ft7">${data.billing.phone}</P></TD>
</TR>
<TR>
	<TD class="tr10 td10"><P class="p3 ft10">&nbsp;</P></TD>
	<TD class="tr10 td11"><P class="p3 ft10">&nbsp;</P></TD>
</TR>
</TABLE>
</div>
<div id="id1_2_2">
<TABLE cellpadding=0 cellspacing=0 class="t2">
<TR>
	<TD class="tr8 td12"></TD>
	<TD class="tr8 td13"><P class="p16 ft7">${new Date(
    data.date_created
  ).toLocaleDateString()}</P></TD>
	<TD class="tr8 td14"><P class="p17 ft11">Date  التاریخ</P></TD>
</TR>
<TR>
	<TD class="tr6 td12"></TD>
	<TD class="tr6 td13"><P class="p18 ft7">${data.currency}</P></TD>
	<TD class="tr6 td14"><P class="p1 ft11">العملة</P></TD>
</TR>
<TR>
	<TD class="tr11 td12"></TD>
	<TD rowspan=2 class="tr12 td13"><P class="p16 ft7">${data.id}</P></TD>
	<TD class="tr11 td14"><P class="p19 ft11">رقم الطلب</P></TD>
</TR>
<TR>
	<TD class="tr9 td12"></TD>
	<TD rowspan=2 class="tr13 td14"><P class="p1 ft11">.Order No</P></TD>
</TR>
<TR>
	<TD class="tr10 td12"></TD>
	<TD class="tr10 td13"><P class="p3 ft10">&nbsp;</P></TD>
</TR>
<TR>
	<TD class="tr11 td12"></TD>
	<TD rowspan=2 class="tr12 td13"><P class="p20 ft7">${
    data.payment_method_title
  }</P></TD>
	<TD class="tr11 td14"><P class="p19 ft11">طریقة الدفع</P></TD>
</TR>
<TR>
	<TD class="tr9 td12"></TD>
	<TD rowspan=2 class="tr13 td14"><P class="p1 ft12">Payment Terms</P></TD>
</TR>
<TR>
	<TD class="tr10 td12"></TD>
	<TD class="tr10 td13"><P class="p3 ft10">&nbsp;</P></TD>
</TR>
</TABLE>
</div>
</div>
<div id="id1_3">
<TABLE cellpadding=0 cellspacing=0 class="t3">
<TR>
	<TD class="tr14 td16"><P class="p22 ft11">رقم المنتج</P></TD>
	<TD class="tr14 td17"><P class="p23 ft11">الإجمالي بعد</P></TD>
	<TD class="tr14 td18"><P class="p22 ft11">قیمة</P></TD>
	<TD class="tr14 td18"><P class="p24 ft11">معدل</P></TD>
	<TD class="tr14 td19"><P class="p23 ft11">الإجمالي قبل</P></TD>
	<TD class="tr14 td20"><P class="p25 ft11">سعر</P></TD>
	<TD class="tr14 td21"><P class="p26 ft11">الكمیة</P></TD>
	<TD class="tr14 td22"><P class="p27 ft11">Item Name اسم الصنف</P></TD>
	<TD class="tr14 td23"><P class="p26 ft11">No</P></TD>
</TR>
<TR>
	<TD class="tr13 td25"><P class="p28 ft11">SKU</P></TD>
	<TD class="tr13 td26"><P class="p17 ft11">الضریبة</P></TD>
	<TD class="tr13 td27"><P class="p17 ft11">الضریبة</P></TD>
	<TD class="tr13 td27"><P class="p17 ft11">الضریبة</P></TD>
	<TD class="tr13 td28"><P class="p17 ft11">الضریبة</P></TD>
	<TD class="tr13 td29"><P class="p17 ft11">الوحدة</P></TD>
	<TD class="tr13 td30"><P class="p29 ft11">Qty</P></TD>
	<TD class="tr13 td32"><P class="p17 ft11">م</P></TD>
</TR>
<TR>
	<TD class="tr13 td26"><P class="p1 ft12">Total after</P></TD>
	<TD class="tr13 td27"><P class="p1 ft12">Vat</P></TD>
	<TD class="tr13 td27"><P class="p1 ft12">Vat</P></TD>
	<TD class="tr13 td28"><P class="p1 ft11">Total</P></TD>
	<TD class="tr13 td29"><P class="p6 ft11">unit</P></TD>
</TR>
<TR>
	<TD class="tr13 td26"><P class="p1 ft12">Vat</P></TD>
	<TD class="tr13 td27"><P class="p1 ft11">Value</P></TD>
	<TD class="tr13 td27"><P class="p17 ft11">Rate</P></TD>
	<TD class="tr13 td28"><P class="p17 ft11">before</P></TD>
	<TD class="tr13 td29"><P class="p30 ft11">price</P></TD>
</TR>
<TR>
	<TD class="tr13 td28"><P class="p17 ft11">VAT</P></TD>
	<TD class="tr13 td29"><P class="p6 ft11">S.R</P></TD>
</TR>
${data.line_items.map(
  (item, index) => `
  <tr>
	<td class="tr14 td25"><p class="p1 ft11">${item.sku}</p></td>
	<td class="tr14 td26"><p class="p17 ft11">${
    +item.subtotal + +item.subtotal_tax
  }</p></td>
	<td class="tr14 td27"><p class="p17 ft11">${item.subtotal_tax}</p></td>
	<td class="tr14 td27"><p class="p17 ft11">${
    data.tax_lines[0].rate_percent + " %"
  }</p></td>
	<td class="tr14 td28"><p class="p17 ft11">${+item.subtotal}</p></td>
	<td class="tr14 td29"><p class="p17 ft11">${(
    +item.subtotal +
    +item.subtotal_tax / +item.quantity
  ).toFixed(0)}</p></td>
	<td class="tr14 td30"><p class="p32 ft11">${item.quantity}</p></td>
	<td class="tr14 td31"><p class="p17 ft12">${item.name}</p></td>
	<td class="tr14 td42"><p class="p5 ft11">${index + 1}</p></td>
</tr>

`
)}

</TABLE>
</div>
<div id="id1_4">
<div id="id1_4_1">
<TABLE cellpadding=0 cellspacing=0 class="t4">
<TR>
	<TD class="tr15 td12"></TD>
	<TD rowspan=2 class="tr7 td44"><P class="p34 ft11">${
    +data.total - +data.total_tax
  }</P></TD>
	<TD class="tr15 td45"><P class="p35 ft11">الإجمالي قبل الضریبة</P></TD>
</TR>
<TR>
	<TD class="tr9 td12"></TD>
	<TD rowspan=2 class="tr13 td45"><P class="p36 ft12">Total before Vat</P></TD>
</TR>
<TR>
<TD class="tr10 td12"></TD>
<TD class="tr10 td44"><P class="p3 ft10">&nbsp;</P></TD>
</TR>
<TR>
<TD class="tr9 td12"></TD>
</TR>
<TR>
	<TD class="tr14 td12"></TD>
	<TD rowspan=2 class="tr16 td44"><P class="p34 ft11">${
    data.discount_total
  }</P></TD>
	<TD class="tr14 td45"><P class="p35 ft12">الخصم الإجمالي</P></TD>
</TR>
<TR>
	<TD class="tr9 td12"></TD>
	<TD rowspan=2 class="tr13 td45"><P class="p36 ft12">Total Discount</P></TD>
</TR>
<TR>
	<TD class="tr14 td12"></TD>
	<TD rowspan=2 class="tr16 td44"><P class="p34 ft11">${data.total_tax}</P></TD>
	<TD class="tr14 td45"><P class="p35 ft11">قیمة الضریبة الإجمالیة</P></TD>
</TR>
<TR>
	<TD class="tr9 td12"></TD>
	<TD rowspan=2 class="tr13 td45"><P class="p36 ft12">Total Vat Value</P></TD>
</TR>
</TABLE>
<P class="p37 ft14"><SPAN class="ft7">${
          data.total
        }</SPAN><SPAN class="ft13">المبلغ المستحق بالریال السعودي Due Balance S.R</SPAN></P>
</div>
<div id="id1_4_2">
<P class="p38 ft11">${data.line_items.reduce(
          (acc, item) => acc + item.bundled_items.length,
          0
        )}</P>
</div>
<div id="id1_4_3">
<P class="p38 ft15">إجمالي الكمیة Total Qty</P>
</div>
</div>
</div>
</body>
</html>

        
        
        `,
        data: {},
        path: `./${data.id}.pdf`,
      };
      pdf
        .create(document, options)
        .then((response) => {
          const invoiceName = `${data.id}.pdf`;
          const invoicePath = path.resolve(invoiceName);
          fs.readFile(invoicePath, (err, data) => {
            if (err) {
              return console.log("err" + err);
            }
          });
        })
        .catch((error) => {
          console.error(error);
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
exports.fetchAllNewOrder = fetchAllNewOrder;
exports.woo_order = woo_order;
// exports.saleReport = saleReport;
// exports.refundReport = refundReport;
exports.downloadSaleFile = downloadSaleFile;
exports.downloadRefundFile = downloadRefundFile;
exports.orderStatusHandler = orderStatusHandler;
exports.changePriceHandler = changePriceHandler;
exports.printBulkInvoice = printBulkInvoice;
