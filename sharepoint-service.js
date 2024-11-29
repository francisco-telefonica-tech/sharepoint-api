// https://telefonicacorp-my.sharepoint.com/my?id=/personal/francisco_jimenezmartin_telefonica_com/Documents/prueba&login_hint=francisco.jimenezmartin@telefonica.com

// invocacion servicio para subir archivo a sharepoint
// this.sharepointService.createFolder(folderName)
//     .then( result => {
//         return res.status(201).json({ message: result });
//     })
//     .catch(err => {
//         return res.status(422).json({ message: err });
//     });

class SharepointService {
  constructor() {
    // this.client = new aws.SecretsManager();
    // this.data = this.client.getSecretValue({ SecretId: '<secret number>' }).promise(); // secret number
    // this.secret = JSON.parse(this.data.SecretString);
    this.axios = require("axios");
  }

  getToken() {
    return axios
      .post(
        "https://accounts.accesscontrol.windows.net/<sharepoint resource id>/tokens/OAuth/2", // sharepoint resource id
        querystring.stringify({
          grant_type: "<client_credentials>", // client credentials
          client_id: "<client_id>", // client ID
          client_secret: this.secret,
          resource: "<ask your sharepoint person>", // sharepoint person
        }),
        {
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
          },
        }
      )
      .then((accessToken) => {
        return axios
          .post(
            "https://telefonicacorp-my.sharepoint.com/my/_api/contextinfo",
            {},
            {
              headers: {
                Authorization: `Bearer ${accessToken}`,
              },
            }
          )
          .then((result) => {
            return result.data.FormDigestValue;
          })
          .catch((error) => {
            return error;
          });
      })
      .catch((error) => {
        return error;
      });
  }

  createFolder(folderName) {
    try {
      const urlSP = `https://telefonicacorp-my.sharepoint.com/my/_api/web/GetFolderByServerRelativeUrl('Documents')/AddFolder(url='${folderName}', overwrite=true)`;

      return this.getToken()
        .then((result) => {
          return axios
            .post(urlSP, folderName, {
              maxBodyLength: Infinity,
              maxContentLength: Infinity,
              headers: {
                Authorization: `Bearer ${token}`, //  ¿TOKEN?
                "X-RequestDigest": result,
              },
            })
            .then(() => {
              return "La carpeta se ha creado correctamente.";
            })
            .catch((error) => {
              return error;
            });
        })
        .catch((error) => {
          return error;
        });
    } catch (error) {
      return error;
    }
  }

  uploadFile(folderName, file) {
    try {
      const urlSP = `https://telefonicacorp-my.sharepoint.com/my/_api/web/GetFolderByServerRelativeUrl('Documents')/${folderName}/Add(url='${file.name}', overwrite=true)`;

      return this.getToken()
        .then((result) => {
          return axios
            .post(urlSP, file, {
              maxBodyLength: Infinity,
              maxContentLength: Infinity,
              headers: {
                Authorization: `Bearer ${token}`, //  ¿TOKEN?
                "X-RequestDigest": result,
              },
            })
            .then(() => {
              return "La carpeta se ha creado correctamente.";
            })
            .catch((error) => {
              return error;
            });
        })
        .catch((error) => {
          return error;
        });
    } catch (error) {
      return error;
    }
  }
}

module.exports = SharepointService;
