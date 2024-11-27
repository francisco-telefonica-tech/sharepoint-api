// https://telefonicacorp-my.sharepoint.com/my?
//         id=%2Fpersonal%2Ffrancisco%5Fjimenezmartin%5Ftelefonica%5Fcom%2FDocuments%2Fprueba
//       & login_hint=francisco%2Ejimenezmartin%40telefonica%2Ecom


class SharepointService {
    constructor() {
        // this.client = new aws.SecretsManager();
        // this.axios = require('axios');
        // this.getRequestDigest = null;
        // this.data = this.client.getSecretValue({ SecretId: '<secret number>' }).promise(); // secret number
    }

    getToken() {
        const secret = JSON.parse(this.data.SecretString);
        return axios.post('https://accounts.accesscontrol.windows.net/<sharepoint resource id>/tokens/OAuth/2', // sharepoint resource id
            querystring.stringify({
                grant_type: '<client_credentials>', // client credentials
                client_id: '<client_id>', // client ID
                client_secret: secret,
                resource: '<ask your sharepoint person>' // sharepoint person
            }), {
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded"
                }
            }
        )
    }

    async createFolder(folderName) {
        try {
            const urlSP = `https://telefonicacorp-my.sharepoint.com/my/_api/web/GetFolderByServerRelativeUrl('Documents')/AddFolder(url='${folderName}', overwrite=true)`;
            
            this.getToken().then( accessToken => {
                this.getRequestDigest = axios.post('https://telefonicacorp-my.sharepoint.com/my/_api/contextinfo', {}, {
                    headers: {
                        "Authorization": `Bearer ${accessToken}`,
        
                    }
                })
            });
            
            await axios.post(urlSP, file, {
                maxBodyLength: Infinity,
                maxContentLength: Infinity,
                headers: {
                            'Authorization': `Bearer ${token}`,
                            'X-RequestDigest': this.getRequestDigest.data.FormDigestValue
                        }
                    }
                )
        } catch (e) {
            return e;
        }
    }

    async uploadFile(folderName, file) {
        try {
            const urlSP = `https://telefonicacorp-my.sharepoint.com/my/_api/web/GetFolderByServerRelativeUrl('Documents')/${folderName}/Add(url='${file.name}', overwrite=true)`;
            
            this.getToken().then( accessToken => {
                this.getRequestDigest = axios.post('https://telefonicacorp-my.sharepoint.com/my/_api/contextinfo', {}, {
                    headers: {
                        "Authorization": `Bearer ${accessToken}`,
        
                    }
                })
            });
            
            await axios.post(urlSP, file, {
                maxBodyLength: Infinity,
                maxContentLength: Infinity,
                headers: {
                            'Authorization': `Bearer ${token}`,
                            'X-RequestDigest': this.getRequestDigest.data.FormDigestValue
                        }
                    }
                )
            return  'Fichero cargado correctamente.';
        } catch (e) {
            return e;
        }
    }

    getUsers () {
        return this.axios.get('https://jsonplaceholder.typicode.com/users')
    }
}

module.exports = SharepointService;