import * as React from 'react'
import * as ReactDOM from 'react-dom'
import { initializeIcons } from '@fluentui/font-icons-mdl2'
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
} from '@fluentui/react/lib/DetailsList'
import {
  PrimaryButton,
  IconButton,
  DefaultButton,
  IButtonProps,
} from '@fluentui/react/lib/Button'
import {
  Icon,
  IIconProps,
  ScrollablePane,
  ScrollbarVisibility,
  Sticky,
  StickyPositionType,
} from '@fluentui/react'
import { SearchBox, ISearchBoxStyles } from '@fluentui/react/lib/SearchBox'
import Pagination from 'office-ui-fabric-react-pagination'
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner'
import { initializeFileTypeIcons } from '@fluentui/react-file-type-icons'
import { useState, useEffect } from 'react'
import {
  Dialog,
  DialogType,
  DialogFooter,
  IDialogFooterStyles,
} from '@fluentui/react/lib/Dialog'
import classes from './NwfDashboard.module.scss'
import { useBoolean } from '@fluentui/react-hooks'
import {
  PeoplePicker,
  PrincipalType,
} from '@pnp/spfx-controls-react/lib/PeoplePicker'
import { loadTheme, createTheme, Theme } from '@fluentui/react'
import '../../../ExternalRef/css/style.css'
import {
  IPersonaSharedProps,
  Persona,
  PersonaSize,
  PersonaPresence,
} from '@fluentui/react/lib/Persona'
import { TextField } from '@material-ui/core'
import { IPersonaProps } from '@fluentui/react/lib/Persona'
import { PoepleSharedUsers } from './PoepleSharedUsers'
interface IColumns {
  Title: string
  ClientName: string
  OrderNo: string
  Assigned: string
  StartDate: string
  EndDate: string
  Status: string
  ID: number
  SharedWith: string | Number[] | any
}

var items: IColumns[] = []
let currentpage: number = 1
var totalPage: number = 30
var filteredItems: IColumns[] = []
var SharedWithUser: IColumns[] = []
var UpdatSWID: IColumns[] = []

var updateItemID: number
const nonwellsFargoTheme = createTheme({
  palette: {
    themePrimary: '#004fa2',
    themeLighterAlt: '#f1f6fb',
    themeLighter: '#cadcf0',
    themeLight: '#9fc0e3',
    themeTertiary: '#508ac8',
    themeSecondary: '#155fae',
    themeDarkAlt: '#004793',
    themeDark: '#003c7c',
    themeDarker: '#002c5b',
    neutralLighterAlt: '#faf9f8',
    neutralLighter: '#f3f2f1',
    neutralLight: '#edebe9',
    neutralQuaternaryAlt: '#e1dfdd',
    neutralQuaternary: '#d0d0d0',
    neutralTertiaryAlt: '#c8c6c4',
    neutralTertiary: '#a19f9d',
    neutralSecondary: '#605e5c',
    neutralPrimaryAlt: '#3b3a39',
    neutralPrimary: '#323130',
    neutralDark: '#201f1e',
    black: '#000000',
    white: '#ffffff',
  },
})
//loadTheme(nonwellsFargoTheme);
const modelProps = {
  isBlocking: true,
  styles: { main: { width: '450px' } },
  topOffsetFixed: true,
}
// const dialogStyles = {main:{width:"450px"}}
const pickerStyles = { root: { width: '450px' } }
const searchBoxStyles: Partial<ISearchBoxStyles> = { root: { width: 300 } }
var allUsers = []
export default function NWDashBoardAdmin(
  INWfDashboardProps,
): React.ReactElement<[]> {
  let siteURL = INWfDashboardProps.context.pageContext.web.absoluteUrl
  let UserEmail = INWfDashboardProps.context.pageContext.user.email
  /*--------------------------------------ButtonIcon---------------------------------*/
  const QuoteIcon = (QuoteID): JSX.Element => (
    <IconButton
      iconProps={{ iconName: 'CustomList' }}
      title="Quote"
      onClick={() => {
        location.href =
          siteURL +
          '/SitePages/GeneralQuoteView.aspx?formID=' +
          QuoteID['data-id']
      }}
      className="QuoteIcon"
    />
  )
  const AddGroup = (propsforicon): JSX.Element => (
    <IconButton
      iconProps={{ iconName: 'AddGroup' }}
      title="AddGroup"
      onClick={() => {
        updateItemID = propsforicon['data-id']
        getuserdetails()
      }}
      className="AddGroup"
    />
  )
  // const modalFooterStyles:Partial <IDialogFooterStyles> = {{margin:"1rem"}}
  const Search = (): JSX.Element => (
    <IconButton
      iconProps={{ iconName: 'Search' }}
      title="SearchIcon"
      className="SearchIcon"
    />
  )
  const addIcon: IIconProps = { iconName: 'Add' }
  // const SearchIcon: IIconProps = { iconName: 'Search'  };
  /*--------------------------------------End of ButtonIcon---------------------------------*/
  // const searchBoxStyles: Partial<ISearchBoxStyles> = { root: { width: 300 } };

  const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(true)
  const [users, setUsers] = useState([])
  const [peopleList, setPeopleList] = React.useState<IPersonaProps[] | any>([])
  var _columns: any = [
    {
      key: 'column1',
      name: 'Client Name',
      fieldName: 'ClientName',
      minWidth: 140,
      maxWidth: 160,
      isResizable: true,
      isRowHeader: true,
      // onColumnClick: test,
      // isSorted: false,
      // isSortedDescending: false,
      // sortAscendingAriaLabel: "Sorted A to Z",
      // sortDescendingAriaLabel: "Sorted Z to A",
      // data: "string",
      // isPadded: true
    },
    {
      key: 'column2',
      name: 'Order No',
      fieldName: 'OrderNo',
      minWidth: 50,
      maxWidth: 90,
      isResizable: true,
      isRowHeader: true,
    },
    /*{
      key: "column3",
      name: "Point Count",
      fieldName: "PointCount",
      minWidth: 50,
      maxWidth: 90,
      isResizable: true
    },

    {
      key: "column4",
      name: "Drivers",
      fieldName: "Drivers",
      minWidth: 50,
      maxWidth: 90,
      isResizable: true
    },*/
    {
      key: 'column5',
      name: 'Assigned To',
      fieldName: 'Assigned',
      minWidth: 120,
      maxWidth: 160,
      isResizable: true,
      isRowHeader: true,
      onRender: (item) =>
        item.Assigned ? (
          <Persona
            imageUrl={
              '/_layouts/15/userphoto.aspx?size=S&username=' +
              item.Assigned.EMail
            }
            text={item.Assigned.Title}
            size={PersonaSize.size32}
          />
        ) : (
          ''
        ),
    },
    {
      key: 'column6',
      name: 'Start Date',
      fieldName: 'StartDate',
      minWidth: 50,
      maxWidth: 90,
      isResizable: true,
      isRowHeader: true,
    },
    {
      key: 'column7',
      name: 'End Date',
      fieldName: 'EndDate',
      minWidth: 50,
      maxWidth: 90,
      isResizable: true,
      isRowHeader: true,
    },
    {
      key: 'column8',
      name: 'Status',
      fieldName: 'Status',
      minWidth: 300,
      maxWidth: 350,
      isResizable: true,
      isRowHeader: true,
    },
    {
      key: 'column9',
      name: 'Quote',
      fieldName: 'Quote',
      minWidth: 50,
      maxWidth: 90,
      isResizable: true,
      isRowHeader: true,
      onRender: (item) =>
        item.Status.toLowerCase() != 'Quoted waiting on PO'.toLowerCase() &&
        item.Status.toLowerCase() !=
          'PO received order entered into production queue'.toLowerCase() ? (
          <Icon
            iconName="PageData"
            styles={{
              root: {
                fontSize: 24,
                fontWeight: 400,
                color: '#004fa2',
                cursor: 'pointer',
              },
            }}
            data-id={item.ID}
            onClick={() => {
              location.href =
                siteURL + '/SitePages/GeneralQuoteView.aspx?formID=' + item.ID
            }}
          />
        ) : (
          'N/A'
        ),
    },
    {
      key: 'column10',
      name: 'Shared With',
      fieldName: 'SharedWith',
      minWidth: 50,
      maxWidth: 90,
      isResizable: true,
      isRowHeader: true,
      onRender: (item) => (
        <Icon
          iconName="AddGroup"
          styles={{
            root: {
              fontSize: 24,
              fontWeight: 400,
              color: '#004fa2',
              cursor: 'pointer',
            },
          }}
          data-id={item.ID}
          onClick={() => {
            updateItemID = item.ID
            getuserdetails()
          }}
        />
      ),
    },
  ]
  const sbWidth = 6
  const sbHeight = 6
  const sbBg = 'gray'
  const sbThumbBg = 'black'

  const csvData = [
    ['firstname', 'lastname', 'email'],
    ['Ahmed', 'Tomi', 'ah@smthing.co.com'],
    ['Raed', 'Labes', 'rl@smthing.co.com'],
    ['Yezzi', 'Min l3b', 'ymin@cocococo.com'],
  ]
  const [Column, setColumn] = useState(true)
  const [rows, setrows] = useState(items) //rows for viewing 5 or 10 data..
  const [masterRows, setmasterRows] = useState(items) //contains all rows..

  useEffect((): void => {
    loadTheme(nonwellsFargoTheme)
    initializeIcons()
    initializeFileTypeIcons(undefined)
    getusergroups()
    getGroupFromList()

    // filterItems();
  }, [])
  return (
    <div style={{ margin: '1rem 2rem' }}>
      {Column && (
        <Spinner
          label="Loading items..."
          size={SpinnerSize.large}
          style={{
            width: '100vw',
            height: '100vh',
            position: 'fixed',
            top: 0,
            left: 0,
            backgroundColor: '#fff',
            zIndex: 10000,
          }}
        />
      )}
      <div className="RequestQuoteAndSearchBox">
        <div style={{ display: 'flex', justifyContent: 'space-between' }}>
          <div>
            <PrimaryButton
              className="RequestQuoteAdmin"
              text="Request Quote"
              href={siteURL + '/SitePages/GeneralQuoteRequest.aspx'}
              iconProps={addIcon}
              style={{ marginRight: '0.5rem' }}
            />
          </div>
          <SearchBox
            className="SearchBox"
            onChange={(e) => searchItems(e.target.value)}
            placeholder="Search by order no."
            disableAnimation
            styles={searchBoxStyles}
          />
        </div>
      </div>
      {masterRows.length > 0 ? (
        <div className="DetailsList">
          <Pagination
            currentPage={currentpage}
            totalPages={
              masterRows.length > 0
                ? Math.ceil(masterRows.length / totalPage)
                : 1
            }
            onChange={(page) => {
              paginate(page)
            }}
          />
          <DetailsList
            columns={_columns}
            items={rows}
            selectionMode={SelectionMode.none}
            //onRenderItemColumn={_renderItemColumn}
          />
          {/* <ScrollablePane
            scrollbarVisibility={ScrollbarVisibility.auto}
            styles={{
              root: {
                selectors: {
                  '.ms-ScrollablePane--contentContainer': {
                    scrollbarWidth: sbWidth,
                    scrollbarColor: `${sbThumbBg} ${sbBg}`,
                  },
                  '.ms-ScrollablePane--contentContainer::-webkit-scrollbar': {
                    width: sbWidth,
                    height: sbHeight,
                  },
                  '.ms-ScrollablePane--contentContainer::-webkit-scrollbar-track': {
                    background: sbBg,
                  },
                  '.ms-ScrollablePane--contentContainer::-webkit-scrollbar-thumb': {
                    background: sbThumbBg,
                  },
                },
              },
            }}
          >
            <DetailsList
              columns={_columns}
              items={rows}
              selectionMode={SelectionMode.none}
              onRenderDetailsHeader={(headerProps, defaultRender) => {
                return (
                  <Sticky
                    stickyPosition={StickyPositionType.Header}
                    isScrollSynced={true}
                    stickyBackgroundColor="transparent"
                  >
                    {defaultRender({
                      ...headerProps,
                      styles: {
                        root: {
                          selectors: {
                            '.ms-DetailsHeader-cellName': {
                              fontWeight: 'bold',
                              fontSize: 13,
                            },
                          },
                          background: '#f5f5f5',
                          borderBottom: '1px solid #ddd',
                          paddingTop: 1,
                        },
                      },
                    })}
                  </Sticky>
                )
              }}
            />
          </ScrollablePane> */}
        </div>
      ) : (
        <div style={{ fontWeight: 'bold', textAlign: 'center' }}>
          No Data Found
        </div>
      )}
      <Dialog
        hidden={hideDialog}
        title="Shared With"
        onDismiss={toggleHideDialog}
        modalProps={modelProps}
      >
        <div
          className="DialogBtn"
          style={{ marginBottom: '2rem', position: 'relative' }}
        >
          <TextField
            style={{
              position: 'absolute',
              height: '0px',
              width: 0,
              border: '0',
              outline: 'none',
            }}
          />
          <PoepleSharedUsers
            peoples={users}
            update={updateSharedWithID}
            GetUserDetails={GetsharedwithUserDetails}
          />
          {/*<PeoplePicker
            context={INWfDashboardProps.context}
            personSelectionLimit={3}
            groupName={""} // Leave this blank in case you want to filter from all users
            showtooltip={true}
            required
            titleText="Shared with"
            //errorMessage={Validation.UserDetailsId}
            showHiddenInUI={false}
            onChange={(e) => SharedWithID(e)}
            principalTypes={[PrincipalType.User]}
            defaultSelectedUsers={users}
            resolveDelay={1000}
            ensureUser={true}
          />*/}
        </div>
        <DialogFooter>
          <PrimaryButton onClick={Save} className="Savebtn" text="Save" />
          <DefaultButton
            onClick={toggleHideDialog}
            className="Cancelbtn"
            text="Cancel"
          />
        </DialogFooter>
      </Dialog>
      <div
        style={{ display: 'flex', justifyContent: 'center', margin: '1rem' }}
      ></div>
    </div>
  )

  function searchItems(keyWord: string): void {
    if (keyWord) {
      var filterdata = items.filter((fItem: IColumns) =>
        fItem.OrderNo
          ? fItem.OrderNo.toLowerCase().indexOf(keyWord.toLowerCase()) != -1
          : '',
      )

      setmasterRows([...filterdata])

      var lastIndex: number = 1 * totalPage
      var firstIndex: number = lastIndex - totalPage
      var paginatedItems: IColumns[] = filterdata.slice(firstIndex, lastIndex)
      currentpage = 1
      setrows([...paginatedItems])
    } else {
      var data = items
      setmasterRows([...data])
      var lastIndex: number = 1 * totalPage
      var firstIndex: number = lastIndex - totalPage
      var paginatedItems: IColumns[] = data.slice(firstIndex, lastIndex)
      currentpage = 1
      setrows([...paginatedItems])
    }
  }

  function paginate(pagenumber): void {
    var lastIndex: number = pagenumber * totalPage
    var firstIndex: number = lastIndex - totalPage
    var paginatedItems: IColumns[] = masterRows.slice(firstIndex, lastIndex)
    currentpage = pagenumber
    setrows([...paginatedItems])
  }
  /*-----------------------------------GETDATA--------------------------------*/
  async function getData(): Promise<void> {
    await INWfDashboardProps.spcontext.lists
      .getByTitle('GeneralQuoteRequestList')
      .items.select(
        'Title',
        'CompanyName',
        'ID',
        'Author/ID',
        'Author/EMail',
        'SharedWith/ID',
        'SharedWith/Title',
        'SharedWith/EMail',
        'UserDetails/Title',
        'UserDetails/EMail',
        'Status',
        'StartDate',
        'EndDate',
        'OrderNo',
      )
      .expand('SharedWith,UserDetails,Author')
      .top(5000)
      .orderBy('Modified', false)
      .get()
      .then(async function (data) {
        items = []
        for (var k = 0; k < data.length; k++) {
          var userBelongsToNWFCompany = allUsers.find(
            (x) => x.ID === data[k].Author.ID,
          ).ID
          var isUsersSharedForthisdata = false
          if (data[k].SharedWith) {
            try {
              isUsersSharedForthisdata = data[k].SharedWith.find(
                (x) => x.EMail === UserEmail,
              ).ID
            } catch (e) {
              console.log(e)
            }
          }
          if (userBelongsToNWFCompany || isUsersSharedForthisdata != false || data[k].Author.EMail == UserEmail) {
            var newitem: IColumns = {
              Title: data[k].Title,
              ID: data[k].ID,
              SharedWith: data[k].SharedWith ? data[k].SharedWith : [],
              ClientName: data[k].CompanyName,
              OrderNo: data[k].OrderNo,
              Assigned: data[k].UserDetails ? data[k].UserDetails[0] : '',
              StartDate: 'N/A',
              EndDate: 'N/A',
              Status: data[k].Status,
            }

            items.push(newitem)
          }
        }

        setmasterRows(items)
        var pagenumber = 1
        var lastIndex: number = pagenumber * totalPage
        var firstIndex: number = lastIndex - totalPage
        //var paginatedItems: IColumns[] = masterRows.slice(firstIndex, lastIndex);//changed for double bind issue
        var paginatedItems: IColumns[] = items.slice(firstIndex, lastIndex)
        currentpage = pagenumber
        setrows(paginatedItems)
        setColumn(false)
        //paginate(1);
      })
      .catch(function (error) {
        setColumn(false)
        //alert('Please contact system admin');
      })
  }

  /*--------------------------------------------------------------------------*/
  async function SharedWithID(event): Promise<void> {
    //console.log(event);
    SharedWithUser = []
    for (let i = 0; i < event.length; i++) {
      await INWfDashboardProps.spcontext.siteUsers
        .getByEmail(event[i].secondaryText)
        .get()
        .then(async function (result): Promise<void> {
          if (result.Id) SharedWithUser.push(result.Id)
        })
        .catch(function (error): void {
          //alert(error);
        })
    }
  }

  async function getuserdetails(): Promise<void> {
    let getSelectedUsers: number[] = []
    let getSharedwithusers = []
    var selectedItem = items.filter((data) => {
      return data['ID'] == updateItemID
    })

    if (selectedItem[0]['SharedWith']) {
      for (var i = 0; i < selectedItem[0]['SharedWith'].length; i++) {
        getSelectedUsers.push(selectedItem[0]['SharedWith'][i].EMail)
        getSharedwithusers.push({
          imageUrl:
            '/_layouts/15/userphoto.aspx?size=S&accountname=' +
            selectedItem[0]['SharedWith'][i].EMail,
          text: selectedItem[0]['SharedWith'][i].Title,
          secondaryText: selectedItem[0]['SharedWith'][i].EMail,
          ID: selectedItem[0]['SharedWith'][i].ID,
          key: i,
          isValid: true,
        })
      }
    }

    //setUsers(getSelectedUsers);
    setUsers(getSharedwithusers)
    toggleHideDialog()
  }

  async function Save(itemid): Promise<void> {
    var requestdata = {
      SharedWithId: { results: SharedWithUser },
    }

    await INWfDashboardProps.spcontext.lists
      .getByTitle('GeneralQuoteRequestList')
      .items.getById(updateItemID)
      .update(requestdata)
      .then(async function (data): Promise<void> {
        //alert("Success");

        await INWfDashboardProps.spcontext.lists
          .getByTitle('GeneralQuoteRequestList')
          .items.getById(updateItemID)
          .select('SharedWith/EMail', 'SharedWith/ID', 'SharedWith/Title')
          .expand('SharedWith')
          .get()
          .then(async function (data): Promise<void> {
            for (var j = 0; j < rows.length; j++) {
              if (rows[j]['ID'] == updateItemID) {
                rows[j]['SharedWith'] = data['SharedWith']
                  ? data['SharedWith']
                  : []
                break
              }
            }

            setrows(rows)
            updateItemID = 0
            toggleHideDialog()
          })
          .catch(function (error) {
            //alert(error);
          })
      })
      .catch(function (error) {
        //alert(error);
      })
  }

  async function updateSharedWithID(event): Promise<void> {
    //console.log(event);
    SharedWithUser = []

    for (let i = 0; i < event.length; i++) {
      await INWfDashboardProps.spcontext.siteUsers
        .getByEmail(event[i].secondaryText)
        .get()
        .then(async function (result): Promise<void> {
          if (result.Id) SharedWithUser.push(result.Id)
        })
        .catch(function (error): void {
          //alert(error);
        })
    }
  }

  function GetsharedwithUserDetails(filterText) {
    var result = peopleList.filter(
      (value, index, self) =>
        index === self.findIndex((t) => t.ID === value.ID),
    )
    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText),
    )
  }

  function doesTextStartWith(text: string, filterText: string): boolean {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0
  }

  async function getusergroups() {
    await INWfDashboardProps.spcontext.siteUsers
      .getByEmail(UserEmail)
      .groups.get()
      .then(async function (data) {
        //console.log(data);
      })
      .catch(function (error) {
        //alert('error');
      })
  }

  async function getGroupFromList() {
    var groups = []
    var usersdata = []
    //await INWfDashboardProps.spcontext.lists.getByTitle("ConfigUsers")
    //.items.select("GroupName/ID","Category").filter("Category eq 'NWF'").expand("GroupName").get().then(async function(data)
    await INWfDashboardProps.spcontext.siteUsers
      .getByEmail(UserEmail)
      .groups.get()
      .then(async function (data) {
        if (data.length > 0) {
          await data.forEach(async (item) => {
            await groups.push(item.Id)
          })

          await groups.forEach(async (groupid) => {
            await INWfDashboardProps.spcontext.siteGroups
              .getById(groupid)
              .users.get()
              .then(async function (result) {
                for (var i = 0; i < result.length; i++) {
                  var userdetails = {
                    key: i,
                    imageUrl:
                      '/_layouts/15/userphoto.aspx?size=S&accountname=' +
                      result[i].Email,
                    text: result[i].Title,
                    secondaryText: result[i].Email,
                    ID: result[i].Id,
                    isValid: true,
                  }

                  await usersdata.push(userdetails)
                }

                allUsers = usersdata
                setPeopleList(usersdata)
                getData()
              })
              .catch(function (err) {
                //alert("Group not found: " + err);
                console.log('Group not found: ' + err)
              })
          })
        } else {
          setPeopleList([])
          getData()
        }
      })
      .catch(function (error) {
        //alert(error)
      })
  }
}
export { NWDashBoardAdmin }

function focusLast():
  | string
  | number
  | boolean
  | {}
  | React.ReactElement<any, string | React.JSXElementConstructor<any>>
  | React.ReactNodeArray
  | React.ReactPortal {
  throw new Error('Function not implemented.')
}
