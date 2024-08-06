import React from 'react'
import React, { useState } from 'react';
import './Button.css';
import { SPHttpClient} from '@microsoft/sp-http';

function API(props) {
  const {context} = props;
  const [data, setData] = useState([]);
  const [list,setlist] = useState('');
  const [object,setobject] = useState({title:''});

    async function getMyData() {
        const token = 'eyJ0eXAiOiJKV1QiLCJub25jZSI6IkRRS0N4QVUxNFI1OU5XX0lFblBuNURxTjNZbUw0S0ZNM0hMUHE0ZEttNVUiLCJhbGciOiJSUzI1NiIsIng1dCI6IktRMnRBY3JFN2xCYVZWR0JtYzVGb2JnZEpvNCIsImtpZCI6IktRMnRBY3JFN2xCYVZWR0JtYzVGb2JnZEpvNCJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8zYzkwYTJmZi02OTFjLTQ4M2EtOGU5NC1mYmNhMWI3ZDRlZGYvIiwiaWF0IjoxNzIyNjg5MjM3LCJuYmYiOjE3MjI2ODkyMzcsImV4cCI6MTcyMjc3NTkzOCwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhYQUFBQXUzQzlPZ2liRVFuSFRuek0ydWc4bWpYYTBZbHRGbk93N0VITVpSNzA0ck91WFpRdEJ2aGJpNE5nT0NKby9ZRHBRU1dBeWVrMFVlTWthdE1maTZWU1phVDhvTDFyQldTd3F3U0tMdjhtdGhzPSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggRXhwbG9yZXIiLCJhcHBpZCI6ImRlOGJjOGI1LWQ5ZjktNDhiMS1hOGFkLWI3NDhkYTcyNTA2NCIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoiU29uaSIsImdpdmVuX25hbWUiOiJTaGlscGEiLCJpZHR5cCI6InVzZXIiLCJpcGFkZHIiOiI0OS40Ny4xMzMuNjIiLCJuYW1lIjoiU2hpbHBhIFNvbmkiLCJvaWQiOiJiZDJhZmMyMy1hYTgwLTQ2NWMtYjliNi00ODFmNDQ3MzE0MTIiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzIwMDIxRDRFMkRFNSIsInJoIjoiMC5BVlFBXzZLUVBCeHBPa2lPbFB2S0czMU8zd01BQUFBQUFBQUF3QUFBQUFBQUFBQ2lBQ0kuIiwic2NwIjoiTWFpbC5TZW5kIG9wZW5pZCBwcm9maWxlIFVzZXIuUmVhZCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6IjJhUmN3eklmVU1oTWE5ZkNKVHhVd05NblRqWXE5NzMyUkxDNGNMVGs2ZmsiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiQVMiLCJ0aWQiOiIzYzkwYTJmZi02OTFjLTQ4M2EtOGU5NC1mYmNhMWI3ZDRlZGYiLCJ1bmlxdWVfbmFtZSI6InNoaWxwYXNAeHJtbGFicy5jb20iLCJ1cG4iOiJzaGlscGFzQHhybWxhYnMuY29tIiwidXRpIjoiYzNwZ2lCNTlUazJlV0lLZnFId0xBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiZjI4YTFmNTAtZjZlNy00NTcxLTgxOGItNmExMmYyYWY2YjZjIiwiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il0sInhtc19jYyI6WyJDUDEiXSwieG1zX2lkcmVsIjoiMSAyNCIsInhtc19zc20iOiIxIiwieG1zX3N0Ijp7InN1YiI6IllhYmtyZFlrX2RQc3FGWTZYRzBIZEVTMkZqM0NxVkRYTGFzQTNjUlZiOGcifSwieG1zX3RjZHQiOjE1MDA1NjkwMDZ9.hhNz_j43bv9CgIghCS72Gc3hlE6MvvojPF6VlU5pMtNBEI93NWUzHF8C8IJmi2yDb21vheOYOmOMZ7yunDVVixn-egnqb1uWueBfmASHvYfTEYEnYrHKfuM02CFSucIN7mxPCrxTGo6dDJwjRMF7v9mlNFc6Jd4iTNTQqIuuD_2hzMB7BNegpmpQBXSqoXfxxJQFAFlNQVWNeZG072IBYl_3sUQvOGz_SiL_WyLwl0vCCX2Ws8GLPvx2w3ZGYdX-HyPsLKs89Vrke77YNjLwe0ckPhysEhYt2vApjtdOTmxgKy1xR0nte1apog1PivQDzKgQU3ipO03MxrrzrxKnqg';
        const response = await fetch('https://graph.microsoft.com/v1.0/me', {
          method: 'GET',
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: 'application/json',
          },
        });
    
        console.log(response);
      }
    async function GetAllList(context) {
        const restApiUrl = context.pageContext.web.absoluteUrl + '/_api/web/lists?select=Title';
        const listTitles = [];
    
        try {
            const response = await context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1);
            const results = await response.json();
    
            results.value.forEach((result) => {
              
                  listTitles.push({ title:result.Title,id:result.id});
          });
            console.log(listTitles);
            setData(listTitles);
        } catch (error) {
            console.error("An error occurred:", error);
            throw error;
        }
    }
    async function createListItem(context, listTitle, formData) {
      const restApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listTitle}')/items`;
      if (!listTitle) {
          return "Please Select a list first";
      }
    
      const body = JSON.stringify({
          "Title": formData.title,
          
          
    
      });
      console.log(body);
    
      const options = {
          headers: {
              Accept: "application/json;odata=nometadata",
              "Content-Type": "application/json;odata=nometadata",
              "odata-version": ""
          },
          body: body
      };
    
      try {
          const response = await context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options);
          if (response.ok) {
              let x = await response.json();
              console.log(x.Id);
              return `List item created with id: ${x.Id}`;
          } else {
              const errorResponse = await response.json();
              throw new Error(`Error creating list item: ${errorResponse.error.message.value}`);
          }
      } catch (error) {
          console.error("An error occurred:", error);
          throw error;
      }
    }
    async function getAllListItems(context, listTitle) {
      const apiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`;
    
      try {
          const response = await context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
          if (!response.ok) {
              throw new Error(`Error fetching list items: ${response.statusText}`);
          }
    
          const data = await response.json();
          console.log(data.value) 
      } catch (error) {
          console.error('Error fetching list items:', error);
          throw error;
      }
    }


  return (
    <>
    <button className='Button' onClick={getMyData}>Fetch API</button>
    <button className="Button" onClick={()=>{GetAllList(context)}}>GetListData </button>
    <button className="Button" onClick={()=>{createListItem(context,'Celebrations',{title:'SharePoint Meetup'})}}>
          create Item
        </button>
        <div>
        <input type="text" value={list} placeholder='Enter List Name' onChange={(event)=>{
          setlist(event.target.value) 
        }} />
       <input type="text" placeholder='Enter title' value={object.title} onChange={(event)=>{
          setobject({title:event.target.value})
        }}/>

        <button className='Button' onClick={()=>{
          createListItem(context,list,object)
        }}>Add Item</button>

      </div>
      <button className='Button'
          onClick={()=>{
            getAllListItems(context,'Celebrations');
          }}
          >GetListItems</button>
     
      
    </>
    
  )
}

export default API