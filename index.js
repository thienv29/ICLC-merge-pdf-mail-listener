const Imap = require("imap");
const fs = require("fs");
const { MailParser } = require("mailparser");
const nodemailer = require("nodemailer");
const PDFMerger = require('pdf-merger-js');

const imapConfig = {
  user: "thienv29@gmail.com",
  password: "ljmi ppjs grrf yjwm",
  host: "imap.gmail.com",
  port: 993,
  tls: true,
  tlsOptions: { rejectUnauthorized: false },
};

const outputPath = "attachments";
const pdfFiles = [];

const imap = new Imap(imapConfig);

function openInbox(cb) {
  imap.openBox("INBOX", true, cb);
}

imap.once("ready", () => {
  openInbox((err, box) => {
    if (err) throw err;

    imap.on("mail", (newMailCount) => {
      console.log(`Bạn có ${newMailCount} email(s) mới`);
      imap.search(["UNSEEN"], (searchError, searchResults) => {
        if (searchError) throw searchError;

        if (searchResults.length === 0) {
          console.log("Không tìm thấy email chưa đọc.");
          return;
        }

        const lastUnseenEmail = searchResults.slice(-1);
        const fetch = imap.fetch(lastUnseenEmail, { bodies: "", struct: true });

        fetch.on("message", (msg, seqno) => {
          console.log("Nhận thông điệp");
          const mailparser = new MailParser();

          msg.on("body", (stream, info) => {
            console.log("Thân email");
            mailparser.on("data", (data) => {
              if (data.type === "attachment") {
                let filePath = `${outputPath}/${data.filename}`;
                let writeStream = fs.createWriteStream(filePath);
                data.content.pipe(writeStream);

                data.content.on("end", () => {
                  data.release();
                });

                writeStream.on("finish", () => {
                  // khi lưu xong tôi cần biết là nó đã lưu đủ chưa
                  console.log("File đã được lưu!");
                  if(filePath.endsWith('.pdf')) {
                    pdfFiles.push(filePath);
                  }
                });
              }
            });

            stream.pipe(mailparser);
          });

          msg.once("end", () => {
            console.log("Kết thúc thông điệp");
          });
        });

        fetch.once("error", (fetchError) => {
          console.log("Lỗi khi nhận thông điệp:", fetchError);
        });

        fetch.once("end", async () => {
          console.log("Xử lý xong email cuối cùng.");
          
          // Merge PDF files after processing all emails
          var merger = new PDFMerger();
          
          for(let filePath of pdfFiles) {
            await merger.add(filePath);
          }
          
          let mergedFilePath = 'merged.pdf';
          await merger.save(mergedFilePath);
          
          console.log('PDF files merged successfully!');

          // Send the merged PDF file as an email attachment
          let transporter = nodemailer.createTransport({
            service: 'gmail',
            auth:{
              user: imapConfig.user,
              pass: imapConfig.password
            } 
          });

          let mailOptions = {
            from: imapConfig.user,
            to: 'thienv29@gmail.com',
            subject: 'Merged PDF File',
            text: 'Here is the merged PDF file.',
            attachments: [
              {
                filename: 'merged.pdf',
                path: mergedFilePath
              }
            ]
          };

          transporter.sendMail(mailOptions, function(error, info){
            if (error) {
              console.log(error);
            } else {
              console.log('Email sent: ' + info.response);

              // Delete the original and merged PDF files after sending the email
              pdfFiles.push(mergedFilePath); // Add the path of the merged file to the array
              
              for(let filePath of pdfFiles) {
                fs.unlinkSync(filePath);
              }
              pdfFiles = []
              
              console.log('Original and merged PDF files deleted successfully!');
            }
          });
        });
      });
    });
  });
});

imap.once("error", (err) => {
  console.log(err);
});

imap.once("end", () => {
  console.log("Kết nối kết thúc");
});

imap.connect();
