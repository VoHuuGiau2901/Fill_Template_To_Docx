const express = require('express');

const cors = require('cors');

const path = require("path");
const Docxtemplater = require("docxtemplater");

const ImageModule = require('docxtemplater-image-module-free');
const fs = require("fs");

const PizZip = require("pizzip");
const libre = require('libreoffice-convert');

const docx_output = path.resolve(__dirname + "/Output/" + "/output.docx");

// Load the docx file as content
const content = fs.readFileSync(path.resolve(__dirname + "/Template/Template.docx"));

let Chatacter_data, Image_data;

const app = express();

app.use(express.json());
app.use(cors());

function fill(req, res) {
    const imageOpts = {
        centered: false,
        getImage: function (tagValue, tagName) {
            return fs.readFileSync(tagValue);
        },
        getSize: function (img, tagValue, tagName) {
            return [150, 150];
        },
    };

    const zip = new PizZip(content);

    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        modules: [new ImageModule(imageOpts)],
    });

    switch (req.params.type) {
        case "UL":
            Chatacter_data = {
                type: "UL",
                CODE: '- Linh động nhan,dễ hòa nhập, chấp nhận những ý kiến mới.\n' +
                    '- Thích nghi, mở cửa cho những ý tưởng mới, sẵn sàng hòa mình vào dòng chảy của cuộc sống.\n' +
                    '- Tính chủ động không cao nếu chưa thích nghi - Bắt chước nhanh.\n' +
                    '- Ôn hòa, thân thiện, cởi mở, lãng mạn.\n' +
                    '- Bộc lộ bản chất theo tâm trạng.\n' +
                    '- Mẫu người cộng đồng xã hội,thích tham gia các hoạt động cộng đồng, từ thiện.',
            };
            Image_data = { image: path.resolve(__dirname, "../Template/UL.jpg") };
            break;
        case "RL":
            Chatacter_data = {
                type: "RL",
                CODE: '- Yêu thích phong cách độc đáo, ấn tượng.\n' +
                    '- Phương pháp khác biệt trong điều hành, quản lý.\n' +
                    '- Khả năng phản biện tốt.\n' +
                    '- Linh động, tính cách mạnh, có lý tưởng riêng.\n' +
                    '- Hứng thú điều huyền bí.\n' +
                    '- Thích suy luận, hay phán đoán, hỏi, đánh giá, lập luận trái ngược.\n' +
                    '- Khả năng kiểm soát công việc vào phút chót.',
            };
            Image_data = { image: path.resolve(__dirname, "../Template/RL.jpg") };
            break;
        case "WC":
        case "WD":
        case "WI":
            Chatacter_data = {
                type: "WC – WD - WI",
                CODE: '- Có tính linh hoạt, dễ thích nghi và thích ứng cao\n' +
                    '- Là người đa mục tiêu, đa kế hoạch, nhìn nhận vấn đề đa chiều, làm nhiều việc cùng lúc.\n' +
                    '- Hứng thú với thử thách, khám phá những điều mới lạ.\n' +
                    '- Thích khám phá bản thân.\n' +
                    '- Là người hướng ngoại, Thích chia sẻ.\n' +
                    '- Nhu cầu đòi hỏi sự tôn trọng và khen ngợi cao.',
            };
            Image_data = { image: path.resolve(__dirname, "../Template/WC_WD_WI.jpg") };
            break;
        case "WP":
        case "WL":
            Chatacter_data = {
                type: "WP - WL",
                CODE: '- Thông thái, biểu tượng của cái đẹp, có sức lôi cuốn tự nhiên.\n' +
                    '- Chú ý phong cách, biểu hiện, vẻ ngoài.\n' +
                    '- Chủ nghĩa hoàn hảo.\n' +
                    '- Khả năng dẫn dắt và tư duy sáng tạo.\n' +
                    '- Nhanh nhẹn, phản ứng nhanh.\n' +
                    '- Làm việc theo cách khác biệt.',
            };
            Image_data = { image: path.resolve(__dirname, "../Template/WP_WL.jpg") };
            break;
        case "WT":
        case "WS":
            Chatacter_data = {
                type: "WT - WS",
                CODE: '- Đại bàng chúa – đại bàng mục tiêu.\n' +
                    '- Lãnh đạo tốt – Thích lãnh đạo người khác.\n' +
                    '- Cái tôi cao,không dễ dàng bị ảnh hưởng.\n' +
                    '- Đề cao quan điểm cá nhân.\n' +
                    '- Lập định rõ ràng và có mục tiêu để hành động.',
            };
            Image_data = { image: path.resolve(__dirname, "../Template/WP_WL.jpg") };
            break;
        case "WE":
            Chatacter_data = {
                type: "WE",
                CODE: '- Quyết định dựa trên cảm xúc.\n' +
                    '- Có khả năng truyền cảm hứng cho cả đội.\n' +
                    '- Cái tôi cao và đề cao quan điểm cá nhân (khả năng lãnh đạo).\n' +
                    '- Là người có tầm nhìn xa.\n' +
                    '- BẠN là mẫu người sống trong thế giới của cảm xúc, cực kì sâu sắc và thích quan tâm đến mọi người.\n' +
                    '- Có nhận thức nhạy bén về cảm xúc nội tâm cũng như cảm xúc của người khác.',
            };
            Image_data = { image: path.resolve(__dirname, "../Template/WP_WL.jpg") };
            break;
    }

    let final_data = { ...Chatacter_data, ...Image_data };

    doc.setData(final_data);

    doc.render();

    const buf = doc.getZip().generate({
        type: "nodebuffer",
        compression: "DEFLATE",
    });

    // Export
    fs.writeFileSync(docx_output, buf);

    // send file to client to download
    if (req.params.option === 'DOCX') {
        res.download(docx_output);
    } else {
        const file = fs.readFileSync(docx_output);

        libre.convert(file, '.pdf', undefined, (err, done) => {
            if (err) {
                console.log(`Error converting file: ${err}`);
            }
            let pdf_output = __dirname + "/Output/" + "/output.pdf";
            fs.writeFileSync(pdf_output, done);
            res.download(pdf_output);

            // delete file on server after sent
            fs.unlink(pdf_output, function (err) {
                if (err) console.log(err);;
            });
        });
    }

    // delete file on server after sent
    fs.unlink(docx_output, function (err) {
        if (err) console.log(err);;
    });
    console.log("sent");
}

app.get("/", (req, res) => {
    res.send("hehe");
});

const port = process.env.PORT || '5000';

app.listen(port,() => {
    console.log(`Server is running ${port}`);
});

