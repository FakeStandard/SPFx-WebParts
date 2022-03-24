import * as React from 'react';
import styles from './SpFxCrud.module.scss';
import { ISpFxCrudProps } from './ISpFxCrudProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export default class SpFxCrud extends React.Component<ISpFxCrudProps, {}> {
  public render(): React.ReactElement<ISpFxCrudProps> {
    return (
      <div className={styles.spFxCrud}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Item ID:</div>
                <input type="text" id='itemId'></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Full Name</div>
                <input type="text" id='fullName'></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Age</div>
                <input type="text" id='age'></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>All Items:</div>
                <div id="allItems"></div>
              </div>
              <div className={styles.buttonSection}>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.createItem}>Create</span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.getItemById}>Read</span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.getAllItems}>Read All</span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.updateItem}>Update</span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.deleteItem}>Delete</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // Create Item
  private createItem = async () => {
    try {
      const addItem = await sp.web.lists.getByTitle('EmployeeDetails').items.add({
        'FullName': document.getElementById('fullName')['value'],
        'Age': document.getElementById('age')['value']
      });

      console.log(`Item create successfully with ID: ${addItem.data.ID}`)
    } catch (err) {
      console.log(err);
    }
  }

  // Get Item by ID
  private getItemById = async () => {
    try {
      const id: number = document.getElementById('itemId')['value'];

      if (id > 0) {
        const item: any = await sp.web.lists.getByTitle('EmployeeDetails').items.getById(id).get();
        document.getElementById('fullName')['value'] = item.FullName;
        document.getElementById('age')['value'] = item.Age;
      } else alert('Please enter a valid item id.')
    } catch (err) {
      console.log(err);
    }
  }

  // Get all items
  private getAllItems = async () => {
    try {
      const items: any[] = await sp.web.lists.getByTitle('EmployeeDetails').items.get();
      console.log(items);
      if (items.length > 0) {
        var html = `<table><tr>
        <th>ID</th>
        <th>Full Name</th>
        <th>Age</th>
        </tr>`

        items.map((item, index) => {
          html += `<tr><td>${item.ID}</td><td>${item.FullName}</td><td>${item.Age}</td></tr>`
        })

        html += '</table>';
        document.getElementById('allItems').innerHTML = html;
      } else alert('List is empty.')
    } catch (err) {
      console.log(err);
    }
  }

  // Update Item
  private updateItem = async () => {
    try {
      const id: number = document.getElementById('itemId')['value'];

      if (id > 0) {
        const itemUpdate = await sp.web.lists.getByTitle('EmployeeDetails').items.getById(id).update({
          'FullName': document.getElementById('fullName')['value'],
          'Age': document.getElementById('age')['value']
        });

        console.log(itemUpdate)
      } else alert('Please enter a valid item id.')

    } catch (err) {
      console.log(err);
    }
  }

  // Delete Item
  private deleteItem = async () => {
    try {
      const id: number = parseInt(document.getElementById('itemId')['value']);

      if (id > 0) {
        let deleteItem = await sp.web.lists.getByTitle('EmployeeDetails').items.getById(id).delete();

        console.log(`Item ID: ${id} deleted successfully.`)
      } else alert('Please enter a valid item id.')
    } catch (err) {
      console.log(err);
    }
  }
}
