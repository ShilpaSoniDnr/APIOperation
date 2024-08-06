import * as React from 'react';
import styles from './ApiOperation.module.scss';
import type { IApiOperationProps } from './IApiOperationProps';
import GraphAPI from './GraphAPI';
//import API from './API'





export default class ApiOperation extends React.Component<IApiOperationProps, {}> {
  public render(): React.ReactElement<IApiOperationProps> {
    

    return (
      <section className={`${styles.apiOperation}`}>
       <p className={`${styles.Primary}`}>API Operation</p>
       {/*<API context={this.props.context}/>*/}
       <GraphAPI/>

      </section>
    );
  }
}
