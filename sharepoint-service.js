// https://telefonicacorp-my.sharepoint.com/my?
//         id=%2Fpersonal%2Ffrancisco%5Fjimenezmartin%5Ftelefonica%5Fcom%2FDocuments%2Fprueba
//       & login_hint=francisco%2Ejimenezmartin%40telefonica%2Ecom

class SharepointService {
    constructor() {
        this.client = new aws.SecretsManager();
    }

    async createFolder(folderName, file, fileName, router) {
        try {
            const urlSP = `https://telefonicacorp-my.sharepoint.com/my/_api/web/GetFolderByServerRelativeUrl('Documents')/${folderName}/Add(url='${fileName}', overwrite=true)`;
            const data = await this.client.getSecretValue({ SecretId: '<secret number>' }).promise(); // secret number
            const secret = JSON.parse(data.SecretString);
            const getToken = await router.post('https://accounts.accesscontrol.windows.net/<sharepoint resource id>/tokens/OAuth/2', // sharepoint resource id
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
            const accessToken = getToken.data.access_token;
            const getRequestDigest = await router.post('https://telefonicacorp-my.sharepoint.com/my/_api/contextinfo', {}, {
                headers: {
                    "Authorization": `Bearer ${accessToken}`,
    
                }
            })
            const formDigestValue = getRequestDigest.data.FormDigestValue;
            await router.post(urlSP, file, {
                maxBodyLength: Infinity,
                maxContentLength: Infinity,
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'X-RequestDigest': formDigestValue
                }
            })
            return  'Fichero cargado correctamente.';
        } catch (e) {
            return e;
        }
    }

    async uploadFile(folderName, file, fileName, router) {
        try {
            const urlSP = `https://telefonicacorp-my.sharepoint.com/my/_api/web/GetFolderByServerRelativeUrl('Documents')/${folderName}/Add(url='${fileName}', overwrite=true)`;
            const data = await this.client.getSecretValue({ SecretId: '<secret number>' }).promise(); // secret number
            const secret = JSON.parse(data.SecretString);
            const getToken = await router.post('https://accounts.accesscontrol.windows.net/<sharepoint resource id>/tokens/OAuth/2', // sharepoint resource id
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
            const accessToken = getToken.data.access_token;
            const getRequestDigest = await router.post('https://telefonicacorp-my.sharepoint.com/my/_api/contextinfo', {}, {
                headers: {
                    "Authorization": `Bearer ${accessToken}`,
    
                }
            })
            const formDigestValue = getRequestDigest.data.FormDigestValue;
            await router.post(urlSP, file, {
                maxBodyLength: Infinity,
                maxContentLength: Infinity,
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'X-RequestDigest': formDigestValue
                }
            })
            return  'Fichero cargado correctamente.';
        } catch (e) {
            return e;
        }
    }
}

module.exports = SharepointService;