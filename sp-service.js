// https://telefonicacorp-my.sharepoint.com/my?id=/personal/francisco_jimenezmartin_telefonica_com/Documents/prueba&login_hint=francisco.jimenezmartin@telefonica.com

class SPService {
  constructor() {
    this.axios = require("axios");
    this.spRequest = require("sp-request");
    this.spr = this.spRequest.create({
      clientId: "xxxxxxxxxxxxxxxxx",
      clientSecret: "xxxxxxxxxxxxxxxxx",
    });
  }

  async createFolder(folder) {
    try {
      const urlSP = `https://telefonicacorp-my.sharepoint.com/my/_api/web/GetFolderByServerRelativeUrl('Documents')/Add/${folder}', overwrite=true)`;

      await this.spr
        .post(urlSP, {
          headers: { Accept: "application/json; odata=verbose" },
        })
        .then((result) => {
          return result;
        })
        .catch((error) => {
          return error;
        });
    } catch ({ message }) {
      return message;
    }
  }

  async uploadFile(folder, file) {
    try {
      const urlSP = `https://telefonicacorp-my.sharepoint.com/my/_api/web/GetFolderByServerRelativeUrl('Documents')/${folder}/Add(url='${file.name}', overwrite=true)`;

      await this.spr
        .post(urlSP, {
          body: file,
          json: false,
          headers: { Accept: "application/json; odata=verbose" },
        })
        .then((result) => {
          return result;
        })
        .catch((error) => {
          return error;
        });
    } catch ({ message }) {
      return message;
    }
  }
}

module.exports = SPService;
