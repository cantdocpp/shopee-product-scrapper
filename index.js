'use strict';

require('events').EventEmitter.defaultMaxListeners = 300;
const cheerio = require('cheerio');
const puppeteer = require('puppeteer');
const { ws, wb } = require('./excel');
require('dotenv').config()

let pages = 0;
const domain = 'https://shopee.co.id';
let links = [];

// Change these categories based on your needs
// You can get all of the category code by visiting your seller center
// Link: https://seller.shopee.co.id/portal/categories
const categories = {
    'Jam Tangan Pria': 12456,
    'Jam Tangan Wanita': 12448,
    'Masker': 18100
};
const excelData = [];

async function app() {
    // You can get your shop id by visiting your store and choose 'semua produk' option
    // You'd be able to see your shop ID in the url
    const shopId = process.env.SHOP_ID

    // Change this URL
    let url = `https://shopee.co.id/shop/${shopId}/search?page=${pages}`;

    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto(url, { waitUntil: 'networkidle2' });
    const html = await page.evaluate(() => document.querySelector('*').outerHTML);

    getAllLinks(html)
}

async function getAllLinks(html) {
    const $ = cheerio.load(html);
    const maxPage = $('.shopee-mini-page-controller__total').text() - 1;
    console.log(maxPage, 'max page');
    await $('a').each((index, element) => {
        const data = $(element).data()
        if (data.sqe) {
            links.push($(element).attr('href'));
        }
    })

    if (pages < maxPage) {
        pages++;
        app();
    } else {
        sendSingleProductHtml(links);
    }
}

async function loopHtmlData(links, page) {
    for (let i = 0; i < links.length; i++) {
        const singleProductUrl = domain + links[i];
        await page.goto(singleProductUrl, { waitUntil: 'networkidle2' });
        const html = await page.evaluate(() => document.querySelector('*').outerHTML);

        await generateData(html);
    }
}

async function sendSingleProductHtml(links) {
    console.log(links.length, 'link length')
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await loopHtmlData(links, page)
    writeExcel();
}

function getImageLink(style) {
    const splitStyle = style.split(';');
    const splitUrl = splitStyle[0].split(':');
    const cleanUrl = splitUrl[2].slice(0, -5);
    const fullUrl = 'https:' + cleanUrl;
    return fullUrl;
}

async function generateData(html) {
    const $ = cheerio.load(html);
    const imgs = [];
    const anchorList = [];
    const stocks = [];
    await $('._2Fw7Qu').each(async (index, element) => {
        const data = $(element).attr('style');
        const imageLink = await getImageLink(data);
        imgs.push(imageLink);
    });

    await $('.JFOy4z').each((index, element) => {
        const anchor = $(element).text();
        anchorList.push(anchor);
    })

    await $('.kIo6pj div').each((index, element) => {
        const text = $(element).text();
        stocks.push(text);
    })

    const categoryName = anchorList[2];
    const categoryCode = categories[categoryName];
    const productName = $('.qaNIZv span').text();
    const productDescription = $('._2u0jt9 span').text();
    const productPrice = $('._3n5NQx').text();
    const productStock = stocks[2];
    const mainImage = imgs[0];
    const productWeight = 500;

    console.log(productName)

    const excelObjectData = {};
    excelObjectData.categoryCode = categoryCode;
    excelObjectData.productName = productName;
    excelObjectData.productDescription = productDescription;
    excelObjectData.productPrice = productPrice;
    excelObjectData.productStock = productStock;
    excelObjectData.mainImage = mainImage;
    excelObjectData.productWeight = productWeight;

    for (let i = 1; i < imgs.length; i++) {
        excelObjectData[`productImage${i}`] = imgs[i];
    }

    excelData.push(excelObjectData);
}

function writeExcel() {
    for (let i = 0; i < excelData.length; i++) {
        if (!excelData[i].productName || !excelData[i].categoryCode) {
            console.log('undefined value found');
        } else {
            ws.cell(i + 1, 1)
                .number(excelData[i].categoryCode)
            ws.cell(i + 1, 2)
                .string(excelData[i].productName)
            ws.cell(i + 1, 3)
                .string(excelData[i].productDescription)
            ws.cell(i + 1, 11)
                .string(excelData[i].productPrice)
            ws.cell(i + 1, 12)
                .string(excelData[i].productStock)
            ws.cell(i + 1, 14)
                .string(excelData[i].mainImage)
            
            if (excelData[i].productImage1) {
                ws.cell(i + 1, 15)
                    .string(excelData[i].productImage1)
            }

            if (excelData[i].productImage2) {
                ws.cell(i + 1, 16)
                    .string(excelData[i].productImage2)
            }

            if (excelData[i].productImage3) {
                ws.cell(i + 1, 17)
                    .string(excelData[i].productImage3)
            }

            if (excelData[i].productImage4) {
                ws.cell(i + 1, 18)
                    .string(excelData[i].productImage4)
            }

            if (excelData[i].productImage5) {
                ws.cell(i + 1, 19)
                    .string(excelData[i].productImage5)
            }

            ws.cell(i + 1, 23)
                .number(excelData[i].productWeight)

            wb.write('upload.xlsx', function(err, stats) {
                if (err) {
                    console.log(err)
                } else {
                    console.log(stats)
                }
            });
        }
    }
}

app();
