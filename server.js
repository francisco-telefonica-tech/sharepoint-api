const express = require("express");
const fileUpload = require("express-fileupload");
const fs = require("fs");
const cors = require("cors");
const SharepointService = require("./sharepoint-service");

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
    this.app.use(express.json());
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
    // ENDPOINT -> CREATE FOLDE IN SHAREPOINT
    this.router.get(
      "/create-folder/:name?",
      (req = express.request, res = express.response) => {
        try {
          if (!req.params.name) {
            return res
              .status(400)
              .json({ msg: "No se ha indicado el nombre de la carpeta." });
          }

          const { name: folderName } = req.params;

          // invocacion servicio para subir archivo a sharepoint
          // this.sharepointService.createFolder(folderName)
          //     .then( result => {
          //         return res.status(201).json({ msg: result });
          //     })
          //     .catch(err => {
          //         return res.status(422).json({ msg: err });
          //     });

          // proceso en servidor
          if (fs.existsSync(folderName)) {
            return res
              .status(400)
              .json({ msg: `La carpeta ${folderName} ya existe.` });
          }

          fs.mkdirSync(folderName, { recursive: true });

          return fs.existsSync(folderName)
            ? res.status(201).json({
                msg: `La carpeta ${folderName} se ha creado correctamente.`,
              })
            : res
                .status(500)
                .json({ msg: `Ha ocurrido un error en el servidor.` });
        } catch (error) {
          res.send({ error: error.code });
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
              .json({ msg: "No se ha seleccionado ningún fichero." });
          }

          if (!req.body.folder) {
            return res
              .status(422)
              .json({ msg: "No se ha indicado la carpeta." });
          }

          const { file } = req.files;
          const { folder: folderName } = req.body;

          // invocacion servicio para subir archivo a sharepoint
          // this.sharepointService.uploadFile(folderName, file)
          //     .then( result => {
          //         return res.status(201).json({ msg: result });
          //     })
          //     .catch(err => {
          //         return res.status(422).json({ msg: err });
          //     });

          // proceso en servidor
          if (fs.existsSync(folderName)) {
            return file.mv(__dirname + "/" + folderName + "/" + file.name)
              ? res.status(201).json({ msg: "Fichero guardado correctamente." })
              : res.status(400).json({
                  msg: "Ha ocurrido un error en el proceso de guardado del fichero.",
                });
          } else {
            return res
              .status(400)
              .json({ msg: "No existe la carpeta indicada." });
          }
        } catch (error) {
          res.send({ error: error.code });
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
