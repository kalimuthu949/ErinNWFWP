import * as React from 'react'
import { useState, useCallback, useRef, useEffect } from 'react'
import {
  BaseClientSideWebPart,
  WebPartContext,
} from '@microsoft/sp-webpart-base'
import { INwfQuoteRequestProps } from './INwfQuoteRequestProps'
import { FontSizes } from '@fluentui/theme'
import { TextField, MaskedTextField } from '@fluentui/react/lib/TextField'
import { Separator } from 'office-ui-fabric-react/lib/Separator'
import { Text, ITextProps } from '@fluentui/react/lib/Text'
import ArrowBackIcon from '@material-ui/icons/ArrowBack'
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from '@fluentui/react/lib/ChoiceGroup'
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button'
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner'
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog'
import {
  PeoplePicker,
  PrincipalType,
} from '@pnp/spfx-controls-react/lib/PeoplePicker'
import {
  Icon,
  loadTheme,
  createTheme,
  Theme,
  ThemeProvider,
  Label,
  classNamesFunction,
} from '@fluentui/react'

// import "office-ui-fabric-react/dist/css/fabric.css";
import styles from './NwfQuoteRequest.module.scss'
import './style.scss'
// import { Hidden } from "@material-ui/core";
import {
  DatePicker,
  IDatePicker,
  mergeStyleSets,
  defaultDatePickerStrings,
} from '@fluentui/react'
// import styles from "./style.scss";
import { PeoplesData } from './PeoplesData'
let getSelectedUsers: number[] = []
let arrAccManager: number[] = []
let NextOrderID: string = ''
const dialogContentProps = {
  type: DialogType.normal,
  title: 'Form Submitted Successfully',
}
var FormFilled: boolean = false
interface formvalues {
  FirstName: string
  LastName: string
  Title: string
  CompanyName: string
  EmailId: string
  PhoneNumber: string
  Address: string
  CompanyWebsite: string
  LyAccMngr: string
  ProposalReq: string
  WorkBegin: string
  WorkComplete: string
  ImpoDate: string
  Description: string
  AdditionalDetails: string
  Scope: string
  SharedLinks: string
  TypesoBid: string
  RemoteConn: boolean
  GraphicsReq: boolean
  NigVersion: boolean
  versionhistory: string
  Speciality: boolean
  UserDetailsId: string | Number[] | any
  Attachments: string[]
}
const wellsFargoTheme = createTheme({
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
const inlineChoiceStyles = {
  flexContainer: {
    display: 'flex',
    label: {
      marginRight: '1rem',
    },
  },
}
const formRowStyle = { display: 'flex', width: '100%', margin: '.5rem' }
const formColumnStyle = { width: '24%', margin: '0rem 0.5rem' }
function NWFQuoteform(props: INwfQuoteRequestProps): React.ReactElement<[]> {
  let siteURL = props.context.pageContext.web.absoluteUrl
  let UserEmail = props.context.pageContext.user.email
  const intialvalues: formvalues[] = [
    {
      FirstName: '',
      LastName: '',
      Title: '',
      CompanyName: '',
      EmailId: '',
      PhoneNumber: '',
      Address: '',
      CompanyWebsite: '',
      LyAccMngr: '',
      ProposalReq: '',
      WorkBegin: '',
      WorkComplete: '',
      ImpoDate: '',
      Description: '',
      AdditionalDetails: '',
      Scope: '',
      SharedLinks: '',
      UserDetailsId: '',
      TypesoBid: '',
      RemoteConn: false,
      GraphicsReq: false,
      NigVersion: false,
      versionhistory: '',
      Speciality: false,
      Attachments: [],
    },
  ]

  const intialvalidations: formvalues[] = [
    {
      FirstName: '',
      LastName: '',
      Title: '',
      CompanyName: '',
      EmailId: '',
      PhoneNumber: '',
      Address: '',
      CompanyWebsite: '',
      LyAccMngr: '',
      ProposalReq: '',
      WorkBegin: '',
      WorkComplete: '',
      ImpoDate: '',
      Description: '',
      AdditionalDetails: '',
      Scope: '',
      SharedLinks: '',
      UserDetailsId: '',
      TypesoBid: '',
      RemoteConn: false,
      GraphicsReq: false,
      NigVersion: false,
      versionhistory: '',
      Speciality: false,
      Attachments: [],
    },
  ]

  const onFormatDate = (date?: Date): string => {
    return date.getMonth() + 1 + '/' + date.getDate() + '/' + date.getFullYear()
  }

  //const arrVlidations:formvalues[]=intialvalidations;

  //const arrValues:formvalues[]=intialvaluestemp;
  const [beginDate, setbeginDate] = useState<Date | undefined>(new Date())
  const [completionDate, setcompletionDate] = useState<Date | undefined>(
    new Date(),
  )
  const [ProposalDate, setProposalDate] = useState<Date | undefined>(new Date())
  const [Column, setColumn] = useState(true)
  const [Hidedialog, setHidedialog] = useState(true)

  ///const[FormFilled,setFormFilled]=useState(false);

  //const[Validation,setValidation]=useState(arrVlidations);
  //const[Submitvalues,setSubmitvalues]=useState(arrValues);

  const [Validation, setValidation] = useState<formvalues[] | undefined>(
    intialvalidations,
  )
  const [Submitvalues, setSubmitvalues] = useState<formvalues[] | undefined>(
    intialvalues,
  )

  const options: IChoiceGroupOption[] = [
    { key: 'A', text: 'Yes' },
    { key: 'B', text: 'No' },
  ]

  let optionsForBid: IChoiceGroupOption[] = [
    { key: '1', text: 'Owner Direct' },
    { key: '2', text: 'dfhd/Competitive' },
  ]

  const [Selectedpeoples, setSelectedpeoples] = useState([])
  const [bidTypes, setBidTypes] = useState(optionsForBid)
  const [niaversion, setniaversion] = useState(false)
  useEffect(() => {
    getLastID()
    getBidTypes()
    setTimeout(() => {
      setColumn(false)
    }, 2000)
  }, [])

  return (
    <ThemeProvider theme={wellsFargoTheme}>
      <div
        id="NWFquoterequest"
        className={styles.nwfQuoteRequest}
        style={{ margin: '1rem 2rem' }}
      >
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
        <div className="ms-Grid" dir="ltr">
          <div style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
            <Icon
              iconName="NavigateBack"
              onClick={DialogBox}
              styles={{
                root: {
                  color: wellsFargoTheme.palette.themePrimary,
                  fontSize: '2rem',
                  cursor: 'pointer',
                },
              }}
            />
            <h2
              style={{
                color: '#004fa2',
                fontWeight: 'bold',
                textAlign: 'center',
                width: '100%',
              }}
            >
              Professional Services Quote
            </h2>
          </div>
          <div className="maincont">
            <div style={formRowStyle}>
              <div style={formColumnStyle}>
                <TextField
                  label="First Name"
                  id="txtDeviceCount"
                  name={'FirstName'}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].FirstName}
                ></TextField>
              </div>
              <div style={formColumnStyle}>
                <TextField
                  label="Last Name"
                  id="txtPointCount"
                  name={'LastName'}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].LastName}
                ></TextField>
              </div>
              <div style={formColumnStyle}>
                <TextField
                  label="Title"
                  id="txtDrivers"
                  name={'Title'}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].Title}
                ></TextField>
              </div>
              <div style={formColumnStyle}>
                <TextField
                  label="Company Name"
                  id="txtSpecialConsiderations"
                  name={'CompanyName'}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].CompanyName}
                />
              </div>
            </div>
            <div style={formRowStyle}>
              <div style={formColumnStyle}>
                <TextField
                  label="Company Website"
                  id="txtManagerEmailID"
                  name={'CompanyWebsite'}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].CompanyWebsite}
                />
              </div>
              <div style={formColumnStyle}>
                <TextField
                  label="E-mail ID"
                  id="txtBEName"
                  name={'EmailId'}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].EmailId}
                />
              </div>
              <div style={formColumnStyle}>
                <TextField
                  label="Phone Number"
                  id="txtBENumber"
                  name={'PhoneNumber'}
                  type="number"
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].PhoneNumber}
                />
              </div>
              <div style={formColumnStyle}>
                <TextField
                  label="Address"
                  id="txtManagerName"
                  multiline
                  rows={3}
                  name={'Address'}
                  resizable={false}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].Address}
                />
              </div>
            </div>
            <div style={formRowStyle}>
              <div style={formColumnStyle}>
                <PeoplePicker
                  context={props.context as any}
                  personSelectionLimit={1}
                  titleText="Lynxspring Account Manager"
                  groupName={'AccountManager'} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  required
                  errorMessage={Validation[0].LyAccMngr}
                  showHiddenInUI={false}
                  onChange={(e) => getAccountmanager(e)}
                  //onChange={this._onItemsChange}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  ensureUser={true}
                />
                {/*<TextField
                        id="txtManagerName"
                        name={"LyAccMngr"}
                        onChange={(e) => handlechange(e)}
                        required
                        errorMessage={Validation[0].LyAccMngr}
                      ></TextField>*/}
              </div>
              <div style={formColumnStyle}>
                {/*<TextField
                  label="Where is Proposal required"
                  id="txtManagerEmailID"
                  name={"ProposalReq"}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].ProposalReq}
                />*/}
                <DatePicker
                  label="When is Proposal required"
                  value={ProposalDate}
                  isRequired
                  onSelectDate={setProposalDate as (date?: Date) => void}
                  formatDate={onFormatDate}
                />
              </div>
              <div style={formColumnStyle}>
                <DatePicker
                  label="When would the work begin ?"
                  value={beginDate}
                  isRequired
                  onSelectDate={setbeginDate as (date?: Date) => void}
                  formatDate={onFormatDate}
                />
                {/* <TextField
                        id="txtVendorManagerName"
                        name={"WorkBegin"}
                        onChange={(e) => handlechange(e)}
                        required
                        errorMessage={Validation[0].WorkBegin}
                      ></TextField>
                    */}
              </div>
              <div style={formColumnStyle}>
                <DatePicker
                  label="When must the work be completed ?"
                  isRequired
                  value={completionDate}
                  onSelectDate={setcompletionDate as (date?: Date) => void}
                  formatDate={onFormatDate}
                />

                {/*                       <TextField
                        id="txtVendorManagerEmailID"
                        name={"WorkComplete"}
                        onChange={(e) => handlechange(e)}
                        required
                        errorMessage={Validation[0].WorkComplete}
                      ></TextField> */}
              </div>
            </div>
            <div style={formRowStyle}>
              <div style={formColumnStyle}>
                <TextField
                  label="Important Dates / Milestones"
                  multiline
                  rows={3}
                  resizable={false}
                  id="txtVendorManagerName"
                  name={'ImpoDate'}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].ImpoDate}
                />
              </div>
              <div style={formColumnStyle}>
                <TextField
                  label="Description of Scope"
                  multiline
                  rows={3}
                  resizable={false}
                  id="txtVendorManagerEmailID"
                  name={'Description'}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].Description}
                />
              </div>
              <div style={formColumnStyle}>
                <TextField
                  label="Additional Details"
                  multiline
                  rows={3}
                  resizable={false}
                  id="txtVendorManagerName"
                  name={'AdditionalDetails'}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].AdditionalDetails}
                />
              </div>
              <div style={formColumnStyle}>
                <Label>Attachment</Label>
                <input
                  className="customfileupload"
                  type="file"
                  onChange={(e) => handleattachemnts(e)}
                  placeholder="Attach full scope of work"
                  id="fileupload"
                />
              </div>
            </div>
            <div style={formRowStyle}>
              {/*<div style={formColumnStyle}>
                <TextField
                  label="Scope"
                  id="txtVendorManagerName"
                  className="scopebox"
                  name={"Scope"}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].Scope}
                />
                    </div>*/}
              <div style={formColumnStyle}>
                <ChoiceGroup
                  label="Types of Bid"
                  name="TypesoBid"
                  defaultSelectedKey={0}
                  options={bidTypes}
                  onChange={_onChange}
                  // label="Pick one"
                  required={true}
                />
              </div>
              <div style={formColumnStyle}>
                <ChoiceGroup
                  styles={inlineChoiceStyles}
                  label="Remote Connectivity"
                  name="RemoteConn"
                  defaultSelectedKey="B"
                  options={options}
                  onChange={_onChange}
                  // label="Pick one"
                  required={true}
                />
              </div>
              <div style={formColumnStyle}>
                <ChoiceGroup
                  styles={inlineChoiceStyles}
                  label="Graphics Required"
                  name="GraphicsReq"
                  defaultSelectedKey="B"
                  options={options}
                  onChange={_onChange}
                  // label="Pick one"
                  required={true}
                />
              </div>
              <div
                style={formColumnStyle}
                className={`niagaraversion ${styles.niagaraversion}`}
              >
                <ChoiceGroup
                  styles={inlineChoiceStyles}
                  label="Niagara Versions"
                  name="NigVersion"
                  defaultSelectedKey="B"
                  options={options}
                  onChange={_onChange}
                  // label="Pick one"
                  required={true}
                />
                {niaversion ? (
                  <TextField
                    styles={{
                      root: {
                        marginTop: '32px',
                        width: '100px',
                      },
                    }}
                    style={{ width: 100 }}
                    id="txtNiagraversion"
                    name={'versionhistory'}
                    onChange={(e) => handlechange(e)}
                    required
                    errorMessage={Validation[0].versionhistory}
                  />
                ) : (
                  ''
                )}
              </div>
            </div>
            <div style={formRowStyle}>
              {/* <ChoiceGroup
              <div style={formColumnStyle}>
                  styles={inlineChoiceStyles}
                  label="Any speciality to built"
                  defaultSelectedKey="B"
                  name="Speciality"
                  options={options}
                  onChange={_onChange}
                  // label="Pick one"
                  required={true}
                />
              </div>*/}
              <div className="peoplepicker" style={formColumnStyle}>
                {/*<PeoplePicker
                  context={props.context as any}
                  titleText="Add User"
                  personSelectionLimit={3}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  required
                  errorMessage={Validation[0].UserDetailsId}
                  showHiddenInUI={false}
                  onChange={(e) => getUserID(e)}
                  //onChange={this._onItemsChange}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  ensureUser={true}
                />*/}
                <PeoplesData
                  useremail={UserEmail}
                  update={UpdateSelectedUsers}
                  spcontext={props.spcontext}
                />
              </div>
              <div style={formColumnStyle}>
                <TextField
                  label="Proj. Docs Shared Links"
                  multiline
                  rows={3}
                  resizable={false}
                  id="txtVendorManagerName"
                  // className="scopebox"
                  name={'SharedLinks'}
                  onChange={(e) => handlechange(e)}
                  required
                  errorMessage={Validation[0].SharedLinks}
                />
              </div>
            </div>
            <div
              style={{
                display: 'flex',
                justifyContent: 'flex-end',
                marginTop: '1rem',
              }}
            >
              <DefaultButton
                text="Cancel"
                style={{ marginRight: '1rem' }}
                onClick={(e) => {
                  //saveattachments();
                  location.href =
                    siteURL + '/SitePages/GeneralRequestDashboard.aspx'
                }}
              />
              <PrimaryButton text="Submit" onClick={mandatoryvalidation} />
            </div>
          </div>

          <Dialog hidden={Hidedialog} dialogContentProps={dialogContentProps}>
            <DialogFooter>
              <PrimaryButton onClick={DialogBox} text="Ok" />
            </DialogFooter>
          </Dialog>
        </div>
      </div>
    </ThemeProvider>
  )

  function handleattachemnts(e): void {
    if (e.target.files.length > 0)
      Submitvalues[0]['Attachments'] = e.target.files
    else Submitvalues[0]['Attachments'] = []

    setSubmitvalues([...Submitvalues])
  }

  function UpdateSelectedUsers(item) {
    console.log('Called parent function')
    var selctedppls = []
    item.forEach(async (element) => {
      await selctedppls.push(element.ID)
    })
    setSelectedpeoples(selctedppls)
  }

  function _onChange(ev: React.FormEvent<HTMLInputElement>, option: any) {
    var name: string = ev.target['attributes'].name.value
    var value: boolean = false

    if (name == 'NigVersion') {
      if (option.text == 'Yes') {
        setniaversion(true)
      } else {
        setniaversion(false)
        Validation[0]['versionhistory'] = ''
        Submitvalues[0]['versionhistory'] = ''
      }
    } else if (name != 'TypesoBid') {
      if (option.text == 'Yes') value = true
      else value = false

      if (value) {
        Validation[0][name] = ''
        Submitvalues[0][name] = value
      } else {
        Submitvalues[0][name] = ''
      }
    } else {
      if (option.text) {
        Validation[0][name] = ''
        Submitvalues[0][name] = option.text
      } else {
        Submitvalues[0][name] = ''
      }
    }

    setValidation([...Validation])
    setSubmitvalues([...Submitvalues])
  }

  function handlechange(e): void {
    var name: string = e.target.attributes.name.value
    var value: string = e.target.value
    if (value) {
      Validation[0][name] = ''
      Submitvalues[0][name] = value
    } else {
      Submitvalues[0][name] = ''
    }

    setSubmitvalues([...Submitvalues])
    setValidation([...Validation])
  }
  function isEmail(email) {
    var regex = /^([a-zA-Z0-9_.+-])+\@(([a-zA-Z0-9-])+\.)+([a-zA-Z0-9]{2,4})+$/
    return regex.test(email)
  }
  /*----------------------------------------mandatoryvalidation--------------------------------------*/
  function mandatoryvalidation(): void {
    var isAllFieldsFilled: boolean = true

    if (!Submitvalues[0].FirstName) {
      Validation[0].FirstName = 'Please Enter First Name'
      isAllFieldsFilled = false
    } else if (!Submitvalues[0].LastName) {
      Validation[0].LastName = 'Please Enter Last Name'
      isAllFieldsFilled = false
    } else if (!Submitvalues[0].Title) {
      Validation[0].Title = 'Please Enter Title'
      isAllFieldsFilled = false
    } else if (!Submitvalues[0].CompanyName) {
      Validation[0].CompanyName = 'Please Enter Company Name'
      isAllFieldsFilled = false
    } else if (!Submitvalues[0].EmailId || !isEmail(Submitvalues[0].EmailId)) {
      Validation[0].EmailId = 'Please Enter Valid Email'
      isAllFieldsFilled = false
    } else if (!Submitvalues[0].PhoneNumber) {
      Validation[0].PhoneNumber = 'Please Enter Phone Number'
      isAllFieldsFilled = false
    } else if (!Submitvalues[0].Address) {
      Validation[0].Address = 'Please Enter Address'
      isAllFieldsFilled = false
    } else if (!Submitvalues[0].CompanyWebsite) {
      Validation[0].CompanyWebsite = 'Please Enter Company Website'
      isAllFieldsFilled = false
    } else if (arrAccManager.length <= 0) {
      Validation[0].LyAccMngr = 'Please Enter Account Manager'
      isAllFieldsFilled = false
    } /*else if (!Submitvalues[0].ProposalReq) {
      Validation[0].ProposalReq = "Please Enter Where is proposal required";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].WorkBegin) {
      Validation[0].WorkBegin = "Please Enter when would the work begin";
      isAllFieldsFilled = false;
    } else if (!Submitvalues[0].WorkComplete) {
      Validation[0].WorkComplete = "Please Enter when must the work be completed";
      isAllFieldsFilled = false;
    }*/ else if (
      !Submitvalues[0].ImpoDate
    ) {
      Validation[0].ImpoDate = 'Please Enter Important dates'
      isAllFieldsFilled = false
    } else if (!Submitvalues[0].Description) {
      Validation[0].Description = 'Please Enter Description'
      isAllFieldsFilled = false
    } else if (!Submitvalues[0].AdditionalDetails) {
      Validation[0].AdditionalDetails = 'Please Enter Additional Details'
      isAllFieldsFilled = false
    } else if (!Submitvalues[0].versionhistory && niaversion) {
      /*else if (!Submitvalues[0].Scope) {
      Validation[0].Scope = "Please Enter scope";
      isAllFieldsFilled = false;
    }*/
      Validation[0].versionhistory = 'Please Enter Version Details'
      isAllFieldsFilled = false
    } else if (!Submitvalues[0].SharedLinks) {
      Validation[0].SharedLinks = 'Please Enter Shared Links'
      isAllFieldsFilled = false
    } /*else if (getSelectedUsers.length <= 0) {
      Validation[0].UserDetailsId = "Please Enter UserDetails";
      isAllFieldsFilled = false;
    }*/

    setValidation([...Validation])

    Submit(isAllFieldsFilled)
  }

  async function Submit(allvaluesfilled): Promise<void> {
    if (allvaluesfilled) {
      await setColumn(true)
      var requestdata = {
        FirstName: Submitvalues[0].FirstName,
        LastName: Submitvalues[0].LastName,
        Title: Submitvalues[0].Title,
        CompanyName: Submitvalues[0].CompanyName,
        EmailID: Submitvalues[0].EmailId,
        PhoneNumber: Submitvalues[0].PhoneNumber,
        Address: Submitvalues[0].Address,
        CompanyWebsite: Submitvalues[0].CompanyWebsite,
        //LynxspringManager: Submitvalues[0].LyAccMngr,
        ProposalRequired: Submitvalues[0].ProposalReq,
        WorkBegin: Submitvalues[0].WorkBegin,
        WorkBeCompleted: Submitvalues[0].WorkComplete,
        ImpDates_x002f_Milestones: Submitvalues[0].ImpoDate,
        DescriptionofScope: Submitvalues[0].Description,
        AdditionalDetails: Submitvalues[0].AdditionalDetails,
        //Scope: Submitvalues[0].Scope,
        SharedLinks: Submitvalues[0].SharedLinks,
        OrderNo: NextOrderID,
        TypeofBid: Submitvalues[0].TypesoBid,
        RemoteConnectivity: Submitvalues[0].RemoteConn,
        GraphicsRequired: Submitvalues[0].GraphicsReq,
        NiagaraVersion: Submitvalues[0].NigVersion,
        Niaversionhistory: niaversion ? Submitvalues[0].versionhistory : '',
        SpecialityToBeBuilt: Submitvalues[0].Speciality,
        Begindate: beginDate,
        CompletionDate: completionDate,
        ProposalDate: ProposalDate,
        //UserDetailsId: { results: getSelectedUsers },
        UserDetailsId: { results: Selectedpeoples },
        AccountmangerId: { results: arrAccManager },
      }
      await props.spcontext.lists
        .getByTitle('GeneralQuoteRequestList')
        .items.add(requestdata)
        .then(async function (data): Promise<void> {
          await saveattachments()
        })
        .catch(function (error): void {
          alert(error)
        })
    }
  }
  function DialogBox(): void {
    location.href = siteURL + '/SitePages/GeneralRequestDashboard.aspx'
  }

  async function saveattachments() {
    var file = Submitvalues[0].Attachments
    if (file.length > 0) {
      if (file[0]['size'] <= 10485760) {
        // small upload
        await props.spcontext
          .getFolderByServerRelativeUrl('ProfessionalQuoteDocuments')
          .files.add(file[0]['name'], file, true)
          .then(function (result) {
            result.file.listItemAllFields.get().then((listItemAllFields) => {
              // get the item id of the file and then update the columns(properties)
              props.spcontext.lists
                .getByTitle('ProfessionalQuoteDocuments')
                .items.getById(listItemAllFields.Id)
                .update({
                  OrderNo: NextOrderID,
                })
                .then((r) => {
                  setColumn(false)
                  setHidedialog(false)
                })
                .catch(function (error) {
                  alert(error)
                })
            })
          })
          .catch(function (error) {
            alert(error)
          })
      } else {
        // large upload
        await props.spcontext
          .getFolderByServerRelativeUrl('ProfessionalQuoteDocuments')
          .files.addChunked(
            file[0]['name'],
            file,
            (data) => {
              console.log({ data: data, message: 'progress' })
            },
            true,
          )
          .then(function (result) {
            result.file.listItemAllFields.get().then((listItemAllFields) => {
              // get the item id of the file and then update the columns(properties)
              props.spcontext.lists
                .getByTitle('ProfessionalQuoteDocuments')
                .items.getById(listItemAllFields.Id)
                .update({
                  OrderNo: NextOrderID,
                })
                .then((r) => {
                  setColumn(false)
                  setHidedialog(false)
                })
                .catch(function (error) {
                  alert(error)
                })
            })
          })
          .catch(function (error) {
            alert(error)
          })
      }
    } else {
      setColumn(false)
      setHidedialog(false)
    }
  }
  async function getUserID(event): Promise<void> {
    Validation[0]['UserDetailsId'] = ''
    setValidation([...Validation])

    if (event.length == 0) {
      getSelectedUsers = []
    }

    for (let i = 0; i < event.length; i++) {
      getSelectedUsers = []
      await props.spcontext.siteUsers
        .getByEmail(event[i].secondaryText)
        .get()
        .then(function (result) {
          if (result.Id) getSelectedUsers.push(result.Id)
        })
        .catch(function (error) {
          console.log(error)
        })
    }
  }

  async function getAccountmanager(event): Promise<void> {
    Validation[0]['LyAccMngr'] = ''
    setValidation([...Validation])

    if (event.length == 0) {
      arrAccManager = []
    }

    for (let i = 0; i < event.length; i++) {
      arrAccManager = []
      await props.spcontext.siteUsers
        .getByEmail(event[i].secondaryText)
        .get()
        .then(function (result) {
          if (result.Id) arrAccManager.push(result.Id)
        })
        .catch(function (error) {
          console.log(error)
        })
    }
  }

  // async function saveFile() {
  //   let formData = new FormData();
  //   formData.append("file", fileupload.files[0]);
  //   await fetch('/upload.php', { method: "POST", body: formData });
  //   alert('the file has been uploaded successfully.');
  // }

  function autoIncrementCustomId(lastRecordId) {
    let increasedNum = Number(lastRecordId.replace('GC-', '')) + 1
    let kmsStr = lastRecordId.substr(0, 3)

    kmsStr = kmsStr + increasedNum.toString()
    console.log(kmsStr)
    NextOrderID = kmsStr
  }

  async function getLastID() {
    await props.spcontext.lists
      .getByTitle('GeneralQuoteRequestList')
      .items.select('ID', 'OrderNo')
      .top(1)
      .orderBy('ID', false)
      .get()
      .then(function (data) {
        if (data.length > 0) {
          autoIncrementCustomId(data[0].OrderNo)
        } else {
          autoIncrementCustomId('GC-0')
        }
      })
      .catch(function (error) {})
  }

  async function getBidTypes() {
    await props.spcontext.lists
      .getByTitle('GeneralQuoteRequestList')
      .fields.filter("EntityPropertyName eq 'TypeofBid'")
      .get()
      .then(function (data) {
        if (data.length > 0) {
          console.log(data)
          optionsForBid = []
          data[0].Choices.map(function (val, key) {
            optionsForBid.push({ key: key, text: val })
          })
          setBidTypes(optionsForBid)
        } else {
        }
      })
      .catch(function (error) {})
  }
}
export { NWFQuoteform }
