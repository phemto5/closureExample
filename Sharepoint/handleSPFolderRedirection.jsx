import React from 'react';
import ReactDOM from 'react-dom';
import $ from 'jquery';
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner'

require('./handleSPFolder.html');
// var message = "Redirecting to Folder"
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
        // console.log('Success Path');
        // console.log(data);
        window.location.href = params.folderUrl;
    })
    .fail((data) => {
        spinner.updateMessage("Folder Not Found");
        // console.log('Error Path');
        // console.log(data);
        params.iterator = 5;
        buildFolderTree(params);
    })

function queryStringtoObject(search) {
    var p = {};
    var q = search.replace('?', '');
    var vars = q.split('&');
    vars.forEach(function (element) {
        var name = element.split('=')[0];
        var value = decodeURI(element.split('=')[1]);
        p[name] = value;
    }, this);

    return p;
}
function entityFolderExists(_url) {
    // console.log('lookging for Folder');
    // console.log(_url);
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
    // console.log('Creating Folder')
    // console.log(_folder);
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

function buildFolderTree(params) {
    processFolder(params);

    function processFolder(params) {
        var url = createRestUrl(params.folderArr.slice(0, params.iterator).join('/'));
        entityFolderExists(url)
            .done((data) => {
                //next
                nextFolder(params);
            })
            .fail((err) => {
                //create folder
                createEntityFolder(params.folderArr.slice(0, params.iterator).join('/'))
                    .done((data) => {
                        nextFolder(params);
                    })
                    .fail((err) => {
                        // console.log('Failed to create folder.')
                        spinner.updateMessage('Epic Fail Could not create folder contact IT Support');
                    })
            });
    }
    function nextFolder(params) {
        params.iterator += 1
        if (params.iterator <= params.folderArr.length) {
            processFolder(params);
        } else {
            window.location.href = params.folderUrl;
        }
    }
}


function createRestUrl(folderString) {
    return 'https://wagstaffinc.sharepoint.com/sites/CRM/_api/Web/GetFolderByServerRelativeUrl(\'' + folderString + '\')';
}