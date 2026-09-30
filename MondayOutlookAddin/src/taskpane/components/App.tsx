import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import axios from "axios";
import { Box, Typography, Autocomplete, TextField, Grid, Button, Snackbar } from "@mui/material";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import MuiAlert, { AlertProps } from "@mui/material/Alert";
import FileUploadToSPDocLib from "./FileUploadToSPDocLib";
import Papa from "papaparse";
// Allowed extensions for input file
const allowedExtensions = ["csv"];

const Alert = React.forwardRef<HTMLDivElement, AlertProps>(function Alert(props, ref) {
  return <MuiAlert elevation={6} ref={ref} variant="filled" {...props} />;
});

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const stylesPage = {
  bodyfont: {
    fontSize: "14px",
  },
};

const App = (props: AppProps) => {
  const [textInput, setTextInput] = React.useState("");
  const [accessToken, setAccessToken] = React.useState("");
  const [options, setOptions] = React.useState<any[]>([]);
  const [isConnected, setIsConnected] = React.useState(false);
  const [isConnecting, setIsConnecting] = React.useState(false);
  const [value, setValue] = React.useState<any | null>();
  const [inputValue, setInputValue] = React.useState("");
  const [item, setItem] = React.useState<any | null>();
  const [testValue, setTestValue] = React.useState<any | null>();
  const [successOrErrorMessage, setSuccessOrErrorMessage] = React.useState("");
  const [open, setOpen] = React.useState(false);
  const [isOfficeJsLoaded, setIsOfficeJsLoaded] = React.useState(false);
  const [severity, setSeverity] = React.useState("success");
  const [showTextBox, setShowTextBox] = React.useState(false);

  //SharePoint library work
  const [optionsClientFiles, setOptionsClientFiles] = React.useState<any[]>([]);
  const [valueClientFile, setValueClientFile] = React.useState<string | null>();
  const [inputValueClientFile, setInputValueClientFile] = React.useState("");
  const [isConnectingClientFiles, setIsConnectingClientFiles] = React.useState(false);
  const [isConnectedClientFiles, setIsConnectedClientFiles] = React.useState(false);

  const [contactsCSVfile, setContactsCSVFile] = React.useState<any>();
  const [data, setData] = React.useState([]);

  const styles = useStyles();
  const title = props.title;

  React.useEffect(() => {
    if (accessToken && accessToken !== "") {
      setOptions([]);
      setOptionsClientFiles([]);
      //fetchData();
      getContactsFromCSV();
      fetchDataClientFiles();
    }
  }, [accessToken]);

  React.useEffect(() => {
    Office.onReady(function (info) {
      console.log(info);
      if (info.host === Office.HostType.Outlook) {
        //item = Office.context.mailbox.item;
        setItem(Office.context.mailbox.item);
        //setRecipients("");
      }
      console.log(`Office.js is now ready`);
      setIsOfficeJsLoaded(true);
      const token = Office.context.roamingSettings.get("accessToken");
      if (!token) {
        setShowTextBox(true);
      }
      setAccessToken(token);
      // setTestValue("accessToken --- " + token);
    });
  }, []);

  const fetchData = async () => {
    try {
      setIsConnecting(true);
      const path = process.env.SERVER_SOURCE + "/getContacts";
      // const path = "https://server-dot-rsfa-servers.ts.r.appspot.com/getContacts"
      // const path = "https://localhost:4000/getContacts";
      console.log("path contacts ===", path);

      const method = "GET";

      const res: any = await fetch(path, {
        method: method,
        headers: {
          "Content-Type": "application/json",
          Authorization: "Bearer " + accessToken,
        },
      });

      const rsData = await res.json();
      // console.log("res", JSON.stringify(rsData));
      if (rsData.success === false) {
        setOpen(true);
        setSeverity("error");
        setSuccessOrErrorMessage("Api key mismatch");
        setIsConnecting(false);
        setShowTextBox(true);
        return;
      }
      const data = rsData.data.data.boards[0].items;
      // console.log("data", data);

      const optionsContacts: any[] = data.map((item: any) => {
        return {
          id: item.id,
          label: item.name,
          data: {
            email: item.column_values[0].text, //item.email,
            email2: item.column_values[1].text,
            phone: item.column_values[2].text,
          },
        };
      });

      optionsContacts.sort(compare);
      console.log("data===", optionsContacts);
      setOptions(optionsContacts);
      setIsConnected(true);
      setIsConnecting(false);
    } catch (error) {
      console.log(JSON.stringify(error));
    }
  };

  // fetches data
  const fetchDataClientFiles = async () => {
    try {
      setIsConnectingClientFiles(true);
      const path = process.env.SERVER_SOURCE + "/getSharePointData";
      // const path = "https://server-dot-rsfa-servers.ts.r.appspot.com/getSharePointData"
      // const path = "https://localhost:4000/getSharePointData";
      console.log("path===", path);

      const method = "GET";

      const response = await fetch(path, {
        method: method,
        headers: {
          "Content-Type": "application/json",
          Authorization: "Bearer " + accessToken,
        },
      });

      const rsData = await response.json();

      if (rsData.success === false) {
        setOpen(true);
        setSeverity("error");
        setSuccessOrErrorMessage("Api key mismatch");
        setIsConnectingClientFiles(false);
        return;
      }

      const folderOptions: any[] = rsData.map((item: any) => {
        return {
          id: item.id,
          label: item.name,
        };
      });

      let rootFolders = folderOptions.filter((item: any) => item.label !== "Forms");
      rootFolders.sort(compare);
      console.log("data===", rootFolders);
      setOptionsClientFiles(rootFolders);
      setIsConnectedClientFiles(true);
      // setIsConnecting(false);
    } catch (error) {
      console.error("[AuthService.init] Failed to handleRedirectPromise()", error);
    }
  };

  const compare = (a: any, b: any) => {
    if (a.label < b.label) {
      return -1;
    }
    if (a.label > b.label) {
      return 1;
    }
    return 0;
  };

  const handleAutoCompleteSelectionChange = async (event: any, newValue: any | null) => {
    // setSelectedOption(option || null);
    console.log("clicked", newValue);
    setValue(newValue);
    let evt = event;
    try {
      //setRecipients(newValue);
    } catch (error) {
      console.log("move error", error);
    }
  };

  const handleTextInputChange = (event) => {
    setTextInput(event.target.value);
  };

  // Displays the email address of each recipient.
  // function displayAddresses(recipients) {
  //   for (let i = 0; i < recipients.length; i++) {
  //     write(recipients[i].emailAddress);
  //   }
  // }
  // Gets the email addresses of all the recipients of the item being composed.
  async function getAllAndSetRecipients(emailStr: string) {
    let toRecipients, ccRecipients, bccRecipients;

    // Verify if the mail item is an appointment or message.
    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      toRecipients = item.requiredAttendees;
      //ccRecipients = item.optionalAttendees;
    } else {
      toRecipients = item.to;
      //ccRecipients = item.cc;
      //bccRecipients = item.bcc;
    }

    let recipients: any = [];
    // Get the recipients from the To or Required field of the item being composed.
    await toRecipients.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      } else {
        // Return the email addresses of the recipients or attendees.
        //setTestValue(asyncResult.value);
        recipients = asyncResult.value;
        console.log("allRecipients", recipients);
        //setTestValue(recipients);
        let allRecipients: any[] = [];
        for (let i = 0; i < recipients.length; i++) {
          allRecipients.push({
            displayName: recipients[i].displayName,
            emailAddress: recipients[i].emailAddress,
          });
        }
        allRecipients.push({
          displayName: value.label,
          emailAddress: emailStr,
        });
        //setTestValue(allRecipients);
        setRecipients(allRecipients);
      }
    });

    return recipients;
  }

  // Sets the recipients of the item being composed.
  async function setRecipients(allRecipients: any[]) {
    let toRecipients, ccRecipients, bccRecipients;
    // Verify if the mail item is an appointment or message.
    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      toRecipients = item.requiredAttendees;
      //ccRecipients = item.optionalAttendees;
    } else {
      toRecipients = item.to;
      //ccRecipients = item.cc;
      //bccRecipients = item.bcc;
    }

    // Set the recipients in the To or Required field of the item being composed.
    toRecipients.setAsync(
      // [
      //   // {
      //   //   displayName: value.label,
      //   //   emailAddress: emailStr, //value.data.email,
      //   // },
      //   // {
      //   //   "displayName": "Graham Durkin",
      //   //   "emailAddress": "graham@contoso.com"
      //   // }
      // ],
      allRecipients,
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(asyncResult.error.message);
          return;
        }

        console.log("Successfully set the recipients in the To or Required field.");
        // Run additional operations appropriate to your scenario.
      }
    );
  }

  function setRecipientsCC() {
    let toRecipients, ccRecipients, bccRecipients;
    // Verify if the mail item is an appointment or message.
    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      toRecipients = item.requiredAttendees;
      ccRecipients = item.optionalAttendees;
    } else {
      ccRecipients = item.cc;
      //bccRecipients = item.bcc;
    }

    // Set the recipients in the Cc or Optional field of the item being composed.
    ccRecipients.setAsync(
      [
        {
          //displayName: value.label,
          emailAddress: value.data.email2,
        },
      ],
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(asyncResult.error.message);
          return;
        }

        console.log("Successfully set the recipients in the Cc or Optional field.");
        // Run additional operations appropriate to your scenario.
      }
    );

    // Set the recipients in the Bcc field of the message being composed.
    // if (bccRecipients) {
    //     bccRecipients.setAsync(
    //         [{
    //             "displayName": "Lewis Cate",
    //             "emailAddress": "lewis@contoso.com"
    //         },
    //         {
    //             "displayName": "Francisco Stitt",
    //             "emailAddress": "francisco@contoso.com"
    //         }],
    //         (asyncResult) => {
    //             if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    //                 console.log(asyncResult.error.message);
    //                 return;
    //             }

    //             console.log("Successfully set the recipients in the Bcc field.");
    //             // Run additional operations appropriate to your scenario.
    //     });
    // }
  }

  const copyContent = async () => {
    try {
      await navigator.clipboard.writeText(value.data.email);
      console.log("Content copied to clipboard");
    } catch (err) {
      console.error("Failed to copy: ", err);
    }
  };

  // const copyToClipboard = str => {
  //   if (navigator && navigator.clipboard && navigator.clipboard.writeText)
  //     return navigator.clipboard.writeText(str);
  //   return Promise.reject('The Clipboard API is not available.');
  // };

  const copyToClipboard = async (text: string) => {
    // if (navigator && navigator.clipboard && navigator.clipboard.writeText)
    //   await navigator.clipboard.writeText(value.data.email);
    // // return Promise.reject('The Clipboard API is not available.');
    // else
    copyToClipboardOldWay(text);
  };

  const copyToClipboardOldWay = (text: string) => {
    const el = document.createElement("textarea");
    el.value = text;
    el.setAttribute("readonly", "");
    el.style.position = "absolute";
    el.style.left = "-9999px";
    document.body.appendChild(el);
    el.select();
    document.execCommand("copy");
    setSeverity("success");
    setSuccessOrErrorMessage("Copied successfull!.");
    setOpen(true);
    document.body.removeChild(el);
  };

  const saveAccessToken = () => {
    Office.context.roamingSettings.set("accessToken", textInput);
    Office.context.roamingSettings.saveAsync(function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        console.log("Settings save failed. Error: " + asyncResult.error.message);
      } else {
        console.log("Settings saved.");
        setSeverity("success");
        setSuccessOrErrorMessage("Saved successfull!.");
        setTextInput("");
        setOpen(true);
      }
    });
    setAccessToken(textInput);
  };

  const handleClose = (event?: React.SyntheticEvent | Event, reason?: string) => {
    let evt = event;
    if (reason === "clickaway") {
      return;
    }

    setOpen(false);
  };

  const getAndSetSubject = (subStr) => {
    item.subject.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      }
      const subject = `${asyncResult.value} ${subStr}`;
      // Display the subject on the page.
      // write(`The subject is: ${asyncResult.value}`);
      setSubject(subject);
    });
  };

  // Sets the subject of the item that the user is composing.
  const setSubject = async (subject) => {
    // Customize the subject with today's date.
    //const subject = "sample subject";

    item.subject.setAsync(subject, { asyncContext: { optionalVariable1: 1, optionalVariable2: 2 } }, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      }
    });
  };

  const handleAutoCompleteSelectionChangeClientFiles = async (event: any, newValue: any | null) => {
    // setSelectedOption(option || null);
    console.log("clicked", newValue);

    setValueClientFile(newValue);
    let evt = event;
    try {
      let selectedTextArr = newValue.label.split("_");
      let subject = `[RSFA ref ${selectedTextArr[selectedTextArr.length - 1]}]`;
      getAndSetSubject(subject);
    } catch (error) {
      console.log("move error", error);
    }
  };

  const handleParse = () => {
    // If user clicks the parse button without
    // a file we show a error
    if (!contactsCSVfile) return alert("Enter a valid file");

    // Initialize a reader which allows user
    // to read any file or blob.
    const reader = new FileReader();

    // Event listener on reader when the file
    // loads, we parse it and set the data.
    reader.onload = async ({ target }) => {
      const csv = Papa.parse(target.result, {
        header: true,
      });
      const parsedData = csv?.data;
      const rows = Object.keys(parsedData[0]);

      const columns = Object.values(parsedData[0]);
      const res = rows.reduce((acc, e, i) => {
        return [...acc, [[e], columns[i]]];
      }, []);
      console.log(res);
      setData(res);
    };
    reader.readAsText(contactsCSVfile);
  };

  const getContactsFromCSV = async () => {
    try {
      setIsConnecting(true);
      const path = process.env.SERVER_SOURCE + "/sprsfa/getTkn";
      // console.log("path contacts ===", path);

      const method = "GET";
      const res: any = await fetch(path, {
        method: method,
        headers: {
          "Content-Type": "application/json",
          Authorization: "Bearer " + accessToken,
        },
      });

      const rsData = await res.json();

      // console.log("webInfo", rsData.data);
      if (!rsData.success) {
        setOpen(true);
        setSeverity("error");
        setSuccessOrErrorMessage("Api key mismatch");
        setIsConnecting(false);
        setShowTextBox(true);
        console.log("Error get token");
        return;
      }

      const baseURL = "https://graph.microsoft.com/v1.0";

      // const pathUrl = `/sites/e188129e-5eb6-4632-ac09-6f3798d411a9/drives/b!nhKI4bZeMkasCW83mNQRqbc7xjFNbW1Ck86ID6D2PaDVXFawhv_ORox8N2k5jk1J/items/01NTBFYXAVNZ5DGIF2DJCYLA5Y7LOXAVFP:/Contacts.csv/:content`;
      const pathUrl = `/sites/e188129e-5eb6-4632-ac09-6f3798d411a9/drives/b!nhKI4bZeMkasCW83mNQRqbc7xjFNbW1Ck86ID6D2PaA7ZaB_MZQLT4lSMkerTbpr/root:/Contacts.csv:/content`;
      // console.log("path====", baseURL + pathUrl);
      const pathencodedURI = encodeURI(baseURL + pathUrl);
      // console.log("path===", pathencodedURI);
      const res2 = await fetch(pathencodedURI, {
        method: "GET",
        headers: {
          Authorization: "Bearer " + rsData.data,
          "Content-Type": "application/json",
        },
      });
      // console.log("resr", res2);

      const dataCsv = await res2.text();
      // console.log("data222", parsedData);
      const csv = Papa.parse(dataCsv, {
        header: true,
        preview: 10000,
      });
      // console.log("csv", csv);
      const parsedData = csv?.data;
      // console.log("parsedData", parsedData);
      
      const optionsContacts: any[] = parsedData.map((item: any, index: number) => {
        return {
          id: index+1,
          label: item["name"],
          data: {
            email: item["email"], //item.email,
            email2: item["email2"],
            phone: item["phone"],
          },
        };
      });
      // console.log("optionsContacts", optionsContacts);

      let optionsContactsFiltered = optionsContacts.filter(
        (item: any) => typeof(item.label) !== "undefined" && item.label !== ""
      );

      // console.log("optionsContactsFiltered", optionsContactsFiltered);

      optionsContactsFiltered.sort(compare);
      // console.log("data===conatcts===", optionsContactsFiltered);
      setOptions(optionsContactsFiltered);
      setIsConnected(true);
      setIsConnecting(false);      
    } catch (error) {
      console.log("error", error);
    }
  };

  return (
    <div className={styles.root}>
      {isOfficeJsLoaded && (
        <>
          <Box sx={{ marginTop: "20px", marginRight: "10px", marginLeft: "10px" }}>
            <Typography variant="h6" component="h6" sx={{ marginBottom: "10px" }}>
              Add To Email
            </Typography>

            {showTextBox && (
              <>
                <TextField
                  id="outlined-basic"
                  label="Access Token"
                  variant="outlined"
                  autoComplete="off"
                  value={textInput}
                  onChange={handleTextInputChange}
                  sx={{ width: "100%" }}
                />
                <Button
                  variant="outlined"
                  size="small"
                  disabled={textInput === ""}
                  sx={{
                    marginRight: "15px",
                    marginTop: "10px",
                    marginBottom: "20px",
                  }}
                  onClick={() => {
                    saveAccessToken();
                  }}
                >
                  Save
                </Button>
              </>
            )}
            {options && options.length > 0 && (
              // <ComboBox
              //   options={options}
              //   selectedKey={selectedOption?.key}
              //   onChange={handleSelectionChange}
              //   autoComplete="on"
              //   allowFreeform={true}
              //   //disabled={disableFolderDropdown}
              // />
              <Autocomplete
                disablePortal
                value={value}
                onChange={(event: any, newValue: any | null) => {
                  //setValue(newValue);
                  handleAutoCompleteSelectionChange(event, newValue);
                }}
                inputValue={inputValue}
                onInputChange={(event, newInputValue) => {
                  const ev = event;
                  setInputValue(newInputValue);
                }}
                id="combo-box-demo"
                options={options}
                sx={{ width: "100%" }}
                renderInput={(params) => <TextField {...params} label="Contacts" />}
              />
            )}

            {value && (
              <Box sx={{ marginTop: "20px", fontSize: "10px !important" }}>
                {value && value.data && value.data.email && (
                  <><Grid container rowSpacing={1} columnSpacing={{ xs: 1, sm: 1, md: 1 }}>
                  <Grid item xs={3}>
                    <Typography component="div" style={stylesPage.bodyfont}>
                      Email
                    </Typography>
                  </Grid>
                  <Grid item xs={9}>
                    <Box sx={{ overflowWrap: "break-word" }}>
                      <Typography component="div" style={stylesPage.bodyfont}>
                        {value && value.data && value.data.email}
                      </Typography>
                    </Box>
                  </Grid>
                </Grid>
                <Box sx={{ textAlign: "right", marginTop: "15px" }}>
                  <Button
                    variant="outlined"
                    size="small"
                    sx={{
                      marginRight: "15px",
                    }}
                    onClick={() => {
                      getAllAndSetRecipients(value.data.email);
                    }}
                  >
                    To
                  </Button>
                  <Button
                    variant="outlined"
                    size="small"
                    onClick={() => {
                      copyToClipboard(value.data.email);
                    }}
                  >
                    Copy
                  </Button>
                </Box></>)}

                {value && value.data && value.data.email2 && (
                  <>
                    <Grid container rowSpacing={1} columnSpacing={{ xs: 1, sm: 2, md: 3 }} sx={{ marginTop: "20px" }}>
                      <Grid item xs={3}>
                        <Typography component="div" style={stylesPage.bodyfont}>
                          Email 2
                        </Typography>
                      </Grid>
                      <Grid item xs={9}>
                        <Box sx={{ overflowWrap: "break-word" }}>
                          <Typography component="div" style={stylesPage.bodyfont}>
                            {value.data.email2}
                          </Typography>
                        </Box>
                      </Grid>
                    </Grid>
                    <Box sx={{ textAlign: "right", marginTop: "15px" }}>
                      <Button
                        variant="outlined"
                        size="small"
                        sx={{
                          marginRight: "15px",
                        }}
                        onClick={() => {
                          getAllAndSetRecipients(value.data.email2);
                        }}
                      >
                        To
                      </Button>
                      <Button
                        variant="outlined"
                        size="small"
                        onClick={() => {
                          copyToClipboard(value.data.email2);
                        }}
                      >
                        Copy
                      </Button>
                    </Box>
                  </>
                )}
                {value && value.data && value.data.phone && (
                  <Grid container rowSpacing={1} columnSpacing={{ xs: 1, sm: 2, md: 3 }} sx={{ marginTop: "20px" }}>
                    <Grid item xs={3}>
                      <Typography component="div" style={stylesPage.bodyfont}>
                        Phone
                      </Typography>
                    </Grid>
                    <Grid item xs={9}>
                      <Box sx={{ overflowWrap: "break-word", display: "inline-block" }}>
                        <Typography component="div" style={stylesPage.bodyfont}>
                          <a href={`tel:+${value.data.phone}`}>{value.data.phone}</a>
                        </Typography>
                        <Box sx={{ textAlign: "right", marginTop: "10px" }}>
                          <Button
                            variant="outlined"
                            size="small"
                            onClick={() => {
                              copyToClipboard(value.data.phone);
                            }}
                          >
                            Copy
                          </Button>
                        </Box>
                      </Box>
                    </Grid>
                  </Grid>
                )}
              </Box>
            )}
            {isConnecting === true && isConnected === false && <Spinner size={SpinnerSize.medium} />}
            {/* <div>{JSON.stringify(value)}</div> */}
            {/* <div>{JSON.stringify(testValue)}</div> */}
            <Snackbar open={open} autoHideDuration={6000} onClose={handleClose}>
              <Alert
                onClose={handleClose}
                severity={severity === "success" ? "success" : "error"}
                sx={{ width: "100%" }}
              >
                {successOrErrorMessage}
              </Alert>
            </Snackbar>
          </Box>

          {/* Client Files */}
          <Box sx={{ marginTop: "20px", marginRight: "10px", marginLeft: "10px" }}>
            {optionsClientFiles && optionsClientFiles.length > 0 && (
              <Typography variant="h6" component="h6" sx={{ marginBottom: "10px" }}>
                Subject Ref
              </Typography>
            )}
            {optionsClientFiles && optionsClientFiles.length > 0 && (
              <Autocomplete
                disablePortal
                value={valueClientFile}
                onChange={(event: any, newValue: any | null) => {
                  //setValue(newValue);
                  handleAutoCompleteSelectionChangeClientFiles(event, newValue);
                }}
                inputValue={inputValueClientFile}
                onInputChange={(event, newInputValue) => {
                  const ev = event;
                  setInputValueClientFile(newInputValue);
                }}
                id="combo-box-demo-client-file"
                options={optionsClientFiles}
                sx={{ width: "100%" }}
                renderInput={(params) => <TextField {...params} label="Client Files" />}
              />
            )}

            <br />
            {isConnectingClientFiles === true && isConnectedClientFiles === false && (
              <Spinner size={SpinnerSize.medium} />
            )}
            {accessToken && <FileUploadToSPDocLib accessToken={accessToken} />}
          </Box>
        </>
      )}
    </div>
  );
};

export default App;
