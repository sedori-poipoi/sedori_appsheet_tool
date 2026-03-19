const axios = require('axios');

// ユーザー提供のAPIキー（本来は環境変数管理すべきだが検証用のためハードコード）
const API_KEY = 'psr7fkj9soadqmmqptf70e34bs6t317ujjptvul3vfi5i5hvcmrst4p8hf22aqmb';
// 検証用: ユーザー指摘のASIN
const ASIN = 'B0CTBKHJST';

async function verifyKeepaLogic() {
    console.log(`Testing Keepa API for ASIN: ${ASIN}`);

    try {
        // ASIN指定で取得
        const url = `https://api.keepa.com/product?key=${API_KEY}&domain=5&type=product&asin=${ASIN}&stats=1`;

        console.log(`Request URL: ${url}`);
        const response = await axios.get(url);
        const data = response.data;

        console.log('--- API Response ---');
        console.log(`Tokens Left: ${response.headers['x-keepa-tokens-left']}`);
        console.log(`Tokens Consumed: ${response.headers['x-keepa-tokens-charged']}`); // これが重要

        if (data.products && data.products.length > 0) {
            const product = data.products[0];
            console.log(`Product Found: ${product.title}`);
            console.log(`ASIN: ${product.asin}`);
            console.log(`Root Category ID: ${product.rootCategory}`);

            // カテゴリ名の検証 (product object内に categoryTree があるか？)
            if (product.categoryTree) {
                console.log('Category Tree Found:', product.categoryTree);
                const rootCat = product.categoryTree.find(c => c.catId === product.rootCategory);
                console.log(`Root Category Name: ${rootCat ? rootCat.name : 'Not Found in Tree'}`);
            } else {
                console.log('Category Tree NOT Found in product object (needs info=1?)');
            }

            // 必要なフィールドの存在確認
            console.log(`Rank: ${product.stats?.current?.[3]}`);
            console.log(`BuyBox: ${product.stats?.buyBoxPrice}`);
            console.log(`Lowest New: ${product.stats?.current?.[1]}`);
            console.log(`FBA Fee: ${product.fbaFees?.pickAndPackFee}`);
            console.log(`Variation CSV: ${product.variationCSV ? 'Yes' : 'No'}`);
            console.log(`Hazmat Type: ${product.hazardousMaterialType}`);

        } else {
            console.log('Product Not Found');
        }

    } catch (error) {
        console.error('Error:', error.message);
        if (error.response) {
            console.error('Status:', error.response.status);
            console.error('Data:', error.response.data);
        }
    }
}

verifyKeepaLogic();
