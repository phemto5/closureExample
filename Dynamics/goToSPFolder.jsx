import React from 'react';
import ReactDOM from 'react-dom';
import $ from 'jquery';
import { Button } from 'office-ui-fabric-react/lib/Button';

require('./goToSPFolder.html');
require('../node_modules/office-ui-fabric-react/dist/css/fabric.css');

var gotoSPFolder = window.gotoSPFolder || {};
gotoSPFolder.WebAPI = gotoSPFolder.WebAPI || {};

getEntity(getURL()).then((data) => {
    const GoButton = () => (<div>
        <Button onClick={goToEntity} >SharePoint Folders for {data.new_name}</Button>
    </div>);

    ReactDOM.render(
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
            parts.folderURL = 'https://wagstaffinc.sharepoint.com/sites/CRM/SitePages/handelSPFolder.aspx?folder=/sites/CRM/account/' + parts.account.name + '/' + parts.query.typename + '/' + parts.entity.new_name;
            window.open(parts.folderURL);
        })
}
function queryStringtoObject(search) {
    var p = {};
    var q = search.substring(1);
    var vars = q.split('&');
    vars.forEach(function (element) {
        var name
        p[element.split('=')[0]] = element.split('=')[1];
    }, this);

    return p;
}
function getURL() {
    var q = queryStringtoObject(window.location.search);
    var win = window.location;
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
    return win.protocol + "//" + win.hostname + "/api/data/v8.1/" + q.typename + "(" + q.id.replace("%7b", "").replace("%7d", "") + ")";
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