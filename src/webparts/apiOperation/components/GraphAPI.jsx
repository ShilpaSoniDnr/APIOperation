import React, { useState, useEffect } from 'react';
import { PublicClientApplication } from '@azure/msal-browser';
import Cookies from 'js-cookie';
import './Button.css';

const msalConfig = {
  auth: {
    clientId: '99f25316-47bf-4c6d-b0ac-4b178de98c42',
    authority: 'https://login.microsoftonline.com/3c90a2ff-691c-483a-8e94-fbca1b7d4edf',
    redirectUri: 'https://xrmlabs.sharepoint.com/sites/xrmlabslive/_layouts/15/workbench.aspx' // Ensure this matches your app's redirect URI
  },
};

const pca = new PublicClientApplication(msalConfig);

const getAccessToken = async () => {
  const request = {
    scopes: ["User.Read", "Sites.ReadWrite.All"]
  };

  try {
    let accessToken = Cookies.get('accessToken');
    console.log('Cached accessToken:', accessToken);

    if (!accessToken) {
      const loginResponse = await pca.loginPopup(request);
      console.log('Login response:', loginResponse);
      accessToken = loginResponse.accessToken;

      // Set cookie with an expiration time (e.g., 1 hour)
      const expirationTime = new Date(new Date().getTime() + 60 * 60 * 1000); // 1 hour
      Cookies.set('accessToken', accessToken, { expires: expirationTime });
    }
    return accessToken;
  } catch (error) {
    console.error('Error getting access token:', error);
    throw error; // Re-throw to handle in the calling function
  }
}

const GraphAPI = (props) => {
  const [data, setData] = useState([]);

  const getList = async () => {
    console.log("You have clicked login");

    try {
      const accessToken = await getAccessToken();
      console.log('Access token:', accessToken);

      const myHeaders = new Headers();
      myHeaders.append("Authorization", `Bearer ${accessToken}`);

      const requestOptions = {
        method: 'GET',
        headers: myHeaders,
        redirect: 'follow'
      };

      const graphResponse = await fetch("https://graph.microsoft.com/v1.0/sites/xrmlabs.sharepoint.com,e5380930-eb71-4161-b333-20dc135c9109,82f02cd9-b117-4525-a3a1-4db73b525b84/lists/f5aad320-b91b-441e-a2a9-7cd0545aeffe/items?expand=fields", requestOptions);

      if (!graphResponse.ok) {
        throw new Error(`HTTP error! Status: ${graphResponse.status}`);
      }

      const result = await graphResponse.json();
      console.log("Graph API Response:", result);
      console.log(result.value);
      setData(result.value);

    } catch (error) {
      console.error('Error authenticating or fetching data:', error);
    }
  };

  return (
    <div>
      <h1>Microsoft GraphAPI</h1>
      <div>
        <button className='Button' onClick={getList}>Get list</button>&nbsp;&nbsp;
        <button className='Button1'>Create list</button>&nbsp;&nbsp;
        <button className='Button2'>Get Item</button>&nbsp;&nbsp;
        <button className='Button3'>Create Item</button>&nbsp;&nbsp;
        <button className='Button4'>Update Item</button>&nbsp;&nbsp;
        <button className='Button5'>Delete Item</button>&nbsp;&nbsp;
      </div>
    </div>
  )
}

export default GraphAPI;
