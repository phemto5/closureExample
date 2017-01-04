import React from 'react';
import ReactDOM from 'react-dom';
import $ from 'jquery';
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner'

require('./handleSPFolder.html');
var gotoSPFolder = window.gotoSPFolder || {};
gotoSPFolder.WebAPI = gotoSPFolder.WebAPI || {};
var params = queryStringtoObject(window.location.search);
params.restUrl = createRestUrl(params.folder);
params.folderUrl = 'https://wagstaffinc.sharepoint.com' + params.folder;
params.folderArr = params.folder.split('/');

class Spin extends React.Component {
    constructor(props) {
        super(props);
        this.state = { message: "Redirecting to Folder" };
        this.updateMessage = this.updateMessage.bind(this);
    }
    updateMessage(msg) {
        this.setState({ message: msg });
    }
    render() {
        return (
            <div >
                <Spinner type={SpinnerType.large} label={this.state.message} />
            </div>
        );
    }
}

var spinner = ReactDOM.render(
    <Spin />,
    document.getElementById('root')
);

entityFolderExists(params.restUrl)
    .done((data) => {
        spinner.updateMessage("Folder Found");
        window.location.href = params.folderUrl;
    })
    .fail((data) => {
        spinner.updateMessage("Folder Not Found");
        params.iterator = 5;
        buildFolderTree(params);
    });

function queryStringtoObject(search) {
    var p = {};
    var q = search.replace('?', '');
    var vars = q.split('&');
    vars.forEach(function (element) {
        var name = element.split('=')[0];
        var value = decodeURI(element.split('=')[1]);
        p[name] = value.replace('.', '');
    }, this);
    return p;
}
function entityFolderExists(_url) {
    spinner.updateMessage('Checking location :' + _url)
    return $.ajax(
        {
            url: _url,
            type: 'GET',
            datatype: 'json',
            header: {
                'X-RequestDigest': $("#__REQUESTDIGEST").val()
            }
        }
    )
}
function createEntityFolder(_folder) {
    spinner.updateMessage('Creating location :' + _folder);
    var _url = 'https://wagstaffinc.sharepoint.com/sites/CRM/_api/Web/folders';
    var _data = JSON.stringify({
        '__metadata': {
            'type': 'SP.Folder'
        },
        'ServerRelativeUrl': _folder
    });
    return $.ajax({
        method: 'POST',
        url: _url,
        data: _data,
        headers: {
            'accept': "application/json;odata=verbose",
            'content-type': "application/json;odata=verbose",
            'X-RequestDigest': $("#__REQUESTDIGEST").val()
        }
    });
}

function createEntityFiles(folder, file) {
    spinner.updateMessage('Creating file :' + folder + '/' + file);
    let _url = 'https://wagstaffinc.sharepoint.com/sites/CRM/_api/web/GetFolderByServerRelativeUrl(\'' + folder + '\')/Files/add(url=\'' + file + '\',overwrite=true)'
    return $.ajax({
        method: 'POST',
        url: _url,
        data: 'Do Not Delete This File',
        headers: {
            'accept': "application/json;odata=verbose",
            'content-type': "application/json;odata=verbose",
            'X-RequestDigest': $("#__REQUESTDIGEST").val()
        }
    });
}
function copyEntityFiles(folder, file) {
    spinner.updateMessage('Coping file :' + decodeURIComponent(folder) + '/' + file);
    var originalFileGuid = '';
    switch (file.split('-')[0]) {
        case 'Notebook':
            // originalFile = 'TemplateNotebook.onetoc2';
            originalFileGuid = '7d6cc7cf-2d9f-4d2c-ba71-2a6d5ef2a9d7';
            break;
        case 'SOP':
            // originalFile = 'SOP.one';
            originalFileGuid = '84091bc0-34d0-438e-a9f7-5ff8bcb50fcd';
            break;
        default:
            break;
    }
    var url = 'https://wagstaffinc.sharepoint.com/sites/CRM/_api/web/getFileById(\'' + originalFileGuid + '\')/copyto(strnewurl=\'' + folder + '/' + file + '\',boverwrite=false)';
    //https://wagstaffinc.sharepoint.com/sites/CRM/_api/web/GetFileByServerRelativeUrl(\'/sites/CRM/Document Templates/TemplateNotebook.onetoc2\')
    return $.ajax({
        method: 'POST',
        url: url,
        datatype: 'json',
        headers: {
            'accept': "application/json;odata=verbose",
            'content-type': "application/json;odata=verbose",
            'X-RequestDigest': $("#__REQUESTDIGEST").val()
        }
    });
}

function buildFolderTree(params) {
    processFolder(params);

    function processFolder(params) {
        var url = createRestUrl(params.folderArr.slice(0, params.iterator).join('/'));
        entityFolderExists(url)
            .done((data) => {
                //next
                nextFolder(params, data);
            })
            .fail((err) => {
                //create folder
                createEntityFolder(params.folderArr.slice(0, params.iterator).join('/'))
                    .done((data) => {
                        nextFolder(params, data);
                    })
                    .fail((err) => {
                        spinner.updateMessage('Epic Fail Could not create folder contact IT Support');
                    })
            });
    }
    function nextFolder(params, data) {
        params.iterator += 1;
        if (data.d) {
            params.folderguid = data.d.UniqueId | null;
        }
        if (params.iterator <= params.folderArr.length) {
            processFolder(params);
        } else {
            // Last Folders
            createEntityFolder(params.folder + '/Project Files')
                .then((data) => {
                    return createEntityFolder(params.folder + '/Time Sheets-Commissioning Reports')
                })
                .then((data) => {
                    return createEntityFiles(params.folder, 'Do_Not_Delete.Save')
                })
                .then((data) => {
                    return copyEntityFiles(params.folder, 'Notebook-' + params.folderArr[6] + '.onetoc2')
                })
                .then((data) => {
                    return copyEntityFiles(params.folder, 'SOP-' + params.folderArr[6] + '.one')
                })
                .then((data) => {
                    window.location.href = params.folderUrl;
                })
        }
    }
}

function createRestUrl(folderString) {
    return 'https://wagstaffinc.sharepoint.com/sites/CRM/_api/Web/GetFolderByServerRelativeUrl(\'' + folderString + '\')';
}
