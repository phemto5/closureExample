import * as React from 'react';
import ReactDOM from 'react-dom';
import $ from 'jquery';
import { Button, ButtonType, IButtonProps, Label } from 'office-ui-fabric-react/lib/index';

require('./goToSPFolder.html');
require('../node_modules/office-ui-fabric-react/dist/css/fabric.css');

var gotoSPFolder = window.gotoSPFolder || {};
gotoSPFolder.WebAPI = gotoSPFolder.WebAPI || {};

getEntity(getURL()).then((data) => {
    class GoButton extends React.Component<IButtonProps, {}> {
        constructor() {
            super();
        }
        render() {
            return (
                <Button
                    onClick={goToEntity}
                    buttonType={ButtonType.compound}
                    description={data.new_name}
                    icon='Folder'>
                    Go To SharePoint Folders
                </Button>
            );
        }
    }


    var goButton = ReactDOM.render(
        <GoButton />,
        document.getElementById('root')
    );
});

function goToEntity() {
    var parts = {};
    parts.query = queryStringtoObject(window.location.search);
    getEntity(getURL())
        .then((data) => {
            parts.entity = data;
            return getAccount(data);
        })
        .then((data) => {
            parts.account = data;
            parts.folderURL = 'https://wagstaffinc.sharepoint.com/sites/CRM/SitePages/handelSPFolder.aspx?folder=/sites/CRM/account/' + parts.account.name + '/' + parts.query.typename + '/' + parts.entity.new_name + '&ProjectID=' + parts.entity.new_projectid;
            window.open(parts.folderURL);
        })
}

function queryStringtoObject(search) {
    var p = {};
    var q = search.replace('?', '');
    var vars = q.split('&');
    vars.forEach(function (element) {
        var name = element.split('=')[0];
        var value = decodeURI(element.split('=')[1]);
        if (value.indexOf('?') > 0) {
            queryStringtoObject(value);
        } else {
            p[name] = value;
        }


    }, this);

    return p;
}
function getURL() {
    var q = queryStringtoObject(window.location.search);
    var win = window.location;
    var url = ''
    switch (q.typename) {
        case 'new_projectactivities': {
            q.typename = 'new_projectactivitieses';
            break;
        }
        case 'account': {
            q.typename = 'accounts';
            break;
        }
        case 'opportunity': {
            q.typename = 'opportunities';
            break;
        }
        default: break;
    }
    url = win.protocol + "//" + win.hostname + "/api/data/v8.1/" + q.typename + "(" + q.id.replace("{", "").replace("}", "") + ")";
    return url;
}

function getEntity(_url) {
    return $.ajax(
        {
            url: _url,
            method: 'GET',
            dataType: 'json',
        }
    );
}
function getAccount(data) {
    var win = window.location;
    var _url = win.protocol + "//" + win.hostname + "/api/data/v8.1/accounts(" + data._new_account_value + ")";
    return $.ajax({
        url: _url,
        method: 'GET',
        dataType: 'json'
    })
}