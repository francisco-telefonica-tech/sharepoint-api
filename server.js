const express = require("express");
const fileUpload = require("express-fileupload");
const fs = require("fs");
const cors = require("cors");
// const SharepointService = require("./sp-service");

class Server {
  constructor() {
    this.app = express();

    this.port = process.env.PORT;
    this.router = express.Router();

    // Middlewares
    this.middelwares();

    // pendiente de implementacion
    // this.sharepointService = new SharepointService();

    // Rutas de la aplicacion
    this.routes();
  }

  middelwares() {
    this.app.use(express.urlencoded({ extended: true }));
    this.app.use(express.send());
    this.app.use("/api", this.router);
    this.app.use(cors());
    this.router.use((req, res, next) => {
      next();
    });
    this.router.use(
      fileUpload({
        limits: { fileSize: 50 * 1024 * 1024 },
        useTempFiles: true,
        tempFileDir: "/tmp/",
      })
    );
  }

  routes() {
    // ENDPOINT -> CREATE FOLDER IN SHAREPOINT
    this.router.get(
      "/create-folder/:folder?",
      (req = express.request, res = express.response) => {
        try {
          if (!req.params.folder) {
            return res
              .status(400)
              .send({ error: "No se ha indicado el nombre de la carpeta." });
          }
          res.status(400).send({
            error: "Falta implementacion crear carpeta.",
          });

          const { folder } = req.params;

          // this.sharepointService
          //   .createFolder(folder)
          //   .then((result) => {
          //     res.status(201).send({
          //       message: result,
          //     });
          //   })
          //   .catch((error) => {
          //     res.status(400).send({
          //       error: error,
          //     });
          //   });
        } catch (error) {
          res.status(400).send({
            error: error.message,
          });
        }
      }
    );

    // ENDPOINT -> UPLOAD FILE TO SHAREPOINT
    this.router.post(
      "/upload-file",
      (req = express.request, res = express.response) => {
        try {
          if (!req.files || !req.files.file) {
            return res
              .status(422)
              .send({ error: "No se ha seleccionado ningún fichero." });
          }

          if (!req.body.folder) {
            return res
              .status(422)
              .send({ error: "No se ha indicado la carpeta." });
          }

          res.status(400).send({
            error: "Falta implementacion carga de fichero.",
          });

          const { file } = req.files;
          const { folder } = req.body;

          // this.sharepointService
          //   .uploadFile(folder, file)
          //   .then((result) => {
          //     res.status(201).send({
          //       message: result,
          //     });
          //   })
          //   .catch((error) => {
          //     res.status(400).send({
          //       message: error,
          //     });
          //   });
        } catch (error) {
          res.status(400).send({
            error: error.message,
          });
        }
      }
    );
  }

  listen() {
    this.app.listen(this.port, () => {
      console.log("Sharepoint Api run in port:", this.port);
    });
  }
}

module.exports = Server;
