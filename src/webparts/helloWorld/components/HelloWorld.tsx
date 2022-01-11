import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { IHelloWorldState } from './IHelloWorldState';
import { TestList } from './TestList';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {
  
  private listItemEntityTypeName: string = undefined;

  constructor(props: IHelloWorldProps, state: IHelloWorldState) {
    super(props);

    this.state = {
      status: this.listNotConfigured(this.props) ? 'Please configure list in Web Part properties' : 'Ready',
      items: []
    };
  }

  public componentWillReceiveProps(nextProps: IHelloWorldProps): void {
    this.listItemEntityTypeName = undefined;
    this.setState({
      status: this.listNotConfigured(nextProps) ? 'Please configure list in Web Part properties' : 'Ready',
      items: []
    });
  }
  
  public render(): React.ReactElement<IHelloWorldProps> {
    const items: JSX.Element[] = this.state.items.map((item: TestList, i: number): JSX.Element => {
      return (
        <li>{item.Title} ({item.Id}) </li>
      );
    });
    
    const disabled: string  = this.listNotConfigured(this.props) ? styles.label : ""  ;
    
    return (
      
      <div className={ styles.helloWorld }>
        <div className={ styles.container }>
              <div className={`${styles.row}`}>
                <div className=''>
                  <p className={`ms-fontColor-black ${ styles.title}` } >
                    Sample SharePoint Operations In React Js
                  </p>
                </div>
              </div>
              <div className={`ms-Grid-row ms-bgColor-themeDark  ${styles.row}`}>
                <div className={`ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1 ${styles.subTitle} orange`}>
                  <a href="#" className={`${styles.button1}`} onClick={() => this.createItem()}>
                    <span className={styles.label}>Create item</span>
                  </a>&nbsp;
                  {/* <a href="#" className={`${styles.button1}`} onClick={() => this.readItem()}>
                    <span className={styles.label}>Read item</span>
                  </a>&nbsp; */}
                  <a href="#" className={`${styles.button1} `} onClick={() => this.readItems()}>
                    <span className={styles.label}>Read all items</span>
                  </a>&nbsp;
                  <a href="#" className={`${styles.button1} `} onClick={() => this.updateItem()}>
                    <span className={styles.label}>Update item</span>
                  </a>&nbsp;
                  <a href="#" className={`${styles.button1} `} onClick={() => this.deleteItem()}>
                    <span className={styles.label}>Delete item</span>
                  </a>
                </div>
              </div>
              <div className={`ms-Grid-row ms-bgColor-themeDark  ${styles.row}`}>
                <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>
                  <ul >
                    <li>
                      {items}
                    </li>
                  </ul>
              
                </div>
              </div>
              <div className={`ms-Grid-row ms-bgColor-themeDark  ${styles.row}`}>
                <div className={`${styles.status}`}>
                  {this.state.status }
                </div>
              </div>
          </div>
      </div>
  

          

    );
  }

  private createItem(): void {
    this.setState({
      status: 'Creating item...',
      items: []
    });

    this.getListItemEntityTypeName()
      .then((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
        const body: string = JSON.stringify({
          '__metadata': {
            'type': listItemEntityTypeName
          },
          'Title': `Item ${new Date()}`
        });
        return this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: body
          });
      })
      .then((response: SPHttpClientResponse): Promise<TestList> => {
        return response.json();
      })
      .then((item: TestList): void => {
        this.setState({
          status: `Item '${item.Title}' (ID: ${item.Id}) successfully created`,
          items: []
        });
      }, (error: any): void => {
        this.setState({
          status: 'Error while creating the item: ' + error,
          items: []
        });
      });
  }

  private readItems(): void {
    this.setState({
      status:'Loading all items...',
      items: []
  });
  this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$select=Title,Id`,
  SPHttpClient.configurations.v1,
  {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'odata-version': ''
    }
  })

  .then((response: SPHttpClientResponse): Promise<{ value: TestList[] }> => {
    return response.json();
  })
  .then((response: { value: TestList[] }): void => {
    this.setState({
      status: `Successfully loaded ${response.value.length} items`,
      items: response.value
    });
  }, (error: any): void => {
    this.setState({
      status: 'Loading all items failed with error: ' + error,
      items: []
    });
  });
}

  private getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$orderby=Id desc&$top=1&$select=id`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { value: { Id: number }[] }): void => {
          if (response.value.length === 0) {
            resolve(-1);
          }
          else {
            resolve(response.value[0].Id);
          }
        });
    });
  }

  private updateItem(): void {
    this.setState({
      status: 'Loading latest items...',
      items: []
    });
    let latestItemId: number = undefined;
    let etag: string = undefined;
    let listItemEntityTypeName: string = undefined;
    this.getListItemEntityTypeName()
      .then((listItemType: string): Promise<number> => {
        listItemEntityTypeName = listItemType;
        return this.getLatestItemId();
      })
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        this.setState({
          status: `Loading information about item ID: ${latestItemId}...`,
          items: []
        });
        return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${latestItemId})?$select=Id`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
      })
      .then((response: SPHttpClientResponse): Promise<TestList> => {
        etag = response.headers.get('ETag');
        return response.json();
      })
      .then((item: TestList): Promise<SPHttpClientResponse> => {
        this.setState({
          status: `Updating item with ID: ${latestItemId}...`,
          items: []
        });
        const body: string = JSON.stringify({
          '__metadata': {
            'type': listItemEntityTypeName
          },
          'Title': `Item ${new Date()}`
        });
        return this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${item.Id})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': '',
              'IF-MATCH': etag,
              'X-HTTP-Method': 'MERGE'
            },
            body: body
          });
      })
      .then((response: SPHttpClientResponse): void => {
        this.setState({
          status: `Item with ID: ${latestItemId} successfully updated`,
          items: []
        });
      }, (error: any): void => {
        this.setState({
          status: `Error updating item: ${error}`,
          items: []
        });
      });
  }

  private deleteItem(): void {
    if (!window.confirm('Are you sure you want to delete the latest item?')) {
      return;
    }

    this.setState({
      status: 'Loading latest items...',
      items: []
    });
    let latestItemId: number = undefined;
    let etag: string = undefined;
    this.getLatestItemId()
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        this.setState({
          status: `Loading information about item ID: ${latestItemId}...`,
          items: []
        });
        return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${latestItemId})?$select=Id`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
      })
      .then((response: SPHttpClientResponse): Promise<TestList> => {
        etag = response.headers.get('ETag');
        return response.json();
      })
      .then((item: TestList): Promise<SPHttpClientResponse> => {
        this.setState({
          status: `Deleting item with ID: ${latestItemId}...`,
          items: []
        });
        return this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${item.Id})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': '',
              'IF-MATCH': etag,
              'X-HTTP-Method': 'DELETE'
            }
          });
      })
      .then((response: SPHttpClientResponse): void => {
        this.setState({
          status: `Item with ID: ${latestItemId} successfully deleted`,
          items: []
        });
      }, (error: any): void => {
        this.setState({
          status: `Error deleting item: ${error}`,
          items: []
        });
      });
  }
   
  private listNotConfigured(props: IHelloWorldProps): boolean {
    return props.listName === undefined ||
      props.listName === null ||
      props.listName.length === 0;
  }


  private getListItemEntityTypeName(): Promise<string> {
    return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
      if (this.listItemEntityTypeName) {
        resolve(this.listItemEntityTypeName);
        return;
      }

      this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { ListItemEntityTypeFullName: string }): void => {
          this.listItemEntityTypeName = response.ListItemEntityTypeFullName;
          resolve(this.listItemEntityTypeName);
        });
    });
  }
}


