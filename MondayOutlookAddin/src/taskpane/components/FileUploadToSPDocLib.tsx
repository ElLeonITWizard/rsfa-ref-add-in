import * as React from "react";
import { Label, Link, makeStyles } from "@fluentui/react-components";
import {
  DefaultButton,
  IStackItemStyles,
  IStackStyles,
  IStackTokens,
  PrimaryButton,
  Separator,
  Spinner,
  SpinnerSize,
  Stack,
  mergeStyles,
} from "@fluentui/react";
import { Snackbar, IconButton, TextField, Autocomplete, Typography, Button, Box } from "@mui/material";
import MuiAlert, { AlertProps } from "@mui/material/Alert";
import HighlightOffIcon from "@mui/icons-material/HighlightOff";
import DragDropFileUpload from "./DragDropFileUpload";

interface IAttachment {
  AttachmentId?: string;
  FileName: string;
  FileBuffer?: ArrayBuffer;
  RedirectUrl?: string;
}

interface AppProps {
  accessToken: string;
}

const Alert = React.forwardRef<HTMLDivElement, AlertProps>(function Alert(props, ref) {
  return <MuiAlert elevation={6} ref={ref} variant="filled" {...props} />;
});

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const stackStyles: IStackStyles = {
  root: {
    //background: DefaultPalette.themeTertiary,
  },
};
const stackItemStyles: IStackItemStyles = {
  root: {
    alignItems: "center",
    //background: DefaultPalette.themePrimary,
    //color: DefaultPalette.white,
    display: "flex",
    height: 50,
    justifyContent: "center",
    width: "100%",
  },
};

// Tokens definition
const stackTokens: IStackTokens = {
  childrenGap: 5,
  padding: 10,
};

const FileUploadToSPDocLib = (props: AppProps) => {
  const [attachments, setAttachments] = React.useState<any[]>([]);
  const [attachmentFile, setAttachmentFile] = React.useState<any>();
  const [isSubmitting, setIsSubmitting] = React.useState<boolean>(false);
  const [open, setOpen] = React.useState(false);
  const [severity, setSeverity] = React.useState("success");
  const [successOrErrorMessage, setSuccessOrErrorMessage] = React.useState("");

  const [accessToken, setAccessToken] = React.useState(props.accessToken);

  //SharePoint library work
  const [optionsClientFiles, setOptionsClientFiles] = React.useState<any[]>([]);
  const [valueClientFile, setValueClientFile] = React.useState<any | null>();
  const [inputValueClientFile, setInputValueClientFile] = React.useState("");
  const [isConnectingClientFiles, setIsConnectingClientFiles] = React.useState(false);
  const [isConnectedClientFiles, setIsConnectedClientFiles] = React.useState(false);

  const [optionsClientFilesSubFolder, setOptionsClientFilesSubFolder] = React.useState<any[]>([]);
  const [valueClientFileSubFolder, setValueClientFileSubFolder] = React.useState<string | null>();
  const [inputValueClientFileSubFolder, setInputValueClientFileSubFolder] = React.useState("");
  const [isConnectingClientFilesSubFolder, setIsConnectingClientFilesSubFolder] = React.useState(false);
  const [isConnectedClientFilesSubFolder, setIsConnectedClientFilesSubFolder] = React.useState(false);

  const [folderId, setFolderId] = React.useState<string | null>();
  const [subFolderName, setSubFolderName] = React.useState<string | null>();

  //Third level dropdown
  const [optionsClientFilesSubSubFolder, setOptionsClientFilesSubSubFolder] = React.useState<any[]>([]);
  const [valueClientFileSubSubFolder, setValueClientFileSubSubFolder] = React.useState<string | null>();
  const [inputValueClientFileSubSubFolder, setInputValueClientFileSubSubFolder] = React.useState("");
  const [isConnectingClientFilesSubSubFolder, setIsConnectingClientFilesSubSubFolder] = React.useState(false);
  const [isConnectedClientFilesSubSubFolder, setIsConnectedClientFilesSubSubFolder] = React.useState(false);
  const [subSubFolderName, setSubSubFolderName] = React.useState<string | null>();

  const [textInput, setTextInput] = React.useState("");

  const inputRef = React.useRef(null);
  const styles = useStyles();

  React.useEffect(() => {
    if (accessToken && accessToken !== "") {
      //setOptions([]);
      setOptionsClientFiles([]);
      fetchDataClientFiles();
    }
  }, [accessToken]);

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
    } catch (error) {
      console.error("[AuthService.init] Failed to handleRedirectPromise()", error);
    }
  };

  //fetch sub folder list
  const fetchDataClientFilesSubFolder = async (selectedFolderId: string) => {
    try {
      setIsConnectedClientFilesSubFolder(false);
      setIsConnectingClientFilesSubFolder(true);
      // const path = process.env.SERVER_SOURCE + "/getSharePointData";
      // const path = "https://server-dot-rsfa-servers.ts.r.appspot.com/getSharePointData"
      // const path = "https://localhost:4000/getSharePointSubFolderData/?folderId=" + selectedFolderId;
      const path = process.env.SERVER_SOURCE + "/getSharePointSubFolderData/?folderId=" + selectedFolderId;
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
        setIsConnectingClientFilesSubFolder(false);
        return;
      }

      const folderOptions: any[] = rsData.map((item: any) => {
        return {
          id: item.id,
          label: item.name,
        };
      });

      let subFolders = folderOptions.filter(
        (item: any) => item.label !== "Forms" && item.label !== "f. Email Communication"
      );
      subFolders.sort(compare);
      console.log("data 222===", subFolders);
      setOptionsClientFilesSubFolder(subFolders);
      setIsConnectedClientFilesSubFolder(true);
      setIsConnectingClientFilesSubFolder(false);
    } catch (error) {
      console.error("[AuthService.init] Failed to handleRedirectPromise()", error);
    }
  };

  //fetch sub folder list
  const fetchDataClientFilesSubSubFolder = async (subFolderName: string) => {
    try {
      setIsConnectedClientFilesSubSubFolder(false);
      setIsConnectingClientFilesSubSubFolder(true);
      // const path = process.env.SERVER_SOURCE + "/getSharePointData";
      // const path = "https://server-dot-rsfa-servers.ts.r.appspot.com/getSharePointData"
      // const path =
      //   "https://localhost:4000/getSharePointSubFolderData/?folderId=" +
      //   folderId +
      //   "&subFolderName=" +
      //   subFolderName;
      const path =
        process.env.SERVER_SOURCE +
        "/getSharePointSubFolderData/?folderId=" +
        folderId +
        "&subFolderName=" +
        subFolderName;

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
        setIsConnectingClientFilesSubSubFolder(false);
        return;
      }

      console.log("rsData");
      console.log(rsData);
      const folderOptions: any[] = rsData.map((item: any) => {
        return {
          id: item.id,
          label: item.name,
        };
      });

      let subFolders = folderOptions.filter(
        (item: any) => item.label !== "Forms" && item.label !== "f. Email Communication"
      );
      subFolders.sort(compare);
      console.log("data 222===", subFolders);
      setOptionsClientFilesSubSubFolder(subFolders);
      setIsConnectedClientFilesSubSubFolder(true);
      setIsConnectingClientFilesSubSubFolder(false);
    } catch (error) {
      console.log("[AuthService.init] Failed to handleRedirectPromise()", error);
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

  const _handleFileSelect = (e: any) => {
    e.preventDefault();
    inputRef.current.click();
  };

  const _renderFilesList = () => {
    console.log("atth", attachments);
    return attachments.map((attachment: IAttachment, index: number) => {
      let itemclass = mergeStyles({
        lineHeight: "32px",
      });
      // let link = attachment.RedirectUrl ? (
      //   <Link className={itemclass} target="_blank" href={attachment.RedirectUrl.replace("1https://", "https://")}>
      //     {attachment.FileName}{" "}
      //   </Link>
      // ) : (
      //   <Label className={itemclass}>{attachment.FileName}</Label>
      // );
      let link = <Label>{attachment.FileName}</Label>;
      return (
        <div key={index}>
          {link}

          <IconButton aria-label="delete" onClick={() => _deleteFileItem(index)} disabled={isSubmitting}>
            <HighlightOffIcon sx={{ color: "red" }} />
          </IconButton>
        </div>
      );
    });
  };

  const _handleFileInputChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    event.persist();
    if (!event.target || !event.target.files) {
      return;
    }
    let attachments = new Array<IAttachment>();

    for (var i = 0; i < event.target.files.length; i++) {
      let file = event.target.files.item(i);
      if (!attachments.some((attachment) => attachment.FileName == file.name)) {
        let contentsAsBuffer: any = await _getFileAsBuffer(file);
        let attachment: IAttachment = {
          FileBuffer: contentsAsBuffer,
          FileName: file.name,
          RedirectUrl: URL.createObjectURL(file),
        };
        attachments.push(attachment);
      }
    }
    // let allAttachments = attachments.concat(attachments);
    console.log("up files 2", attachments);

    setAttachments(attachments);
    setAttachmentFile(event.target.files[0]);
    //this.props.valueChanged(allAttachments);
    //Clear Input field
    inputRef.current.value = "";
  };

  const _handleFileInputChangeDropZone = async (file: any) => {
    let attachments = new Array<IAttachment>();

    if (!attachments.some((attachment) => attachment.FileName == file.name)) {
      let contentsAsBuffer: any = await _getFileAsBuffer(file);
      let attachment: IAttachment = {
        FileBuffer: contentsAsBuffer,
        FileName: file.name,
        RedirectUrl: URL.createObjectURL(file),
      };
      attachments.push(attachment);
    }

    // let allAttachments = attachments.concat(attachments);
    console.log("up files 222", attachments);

    setAttachments(attachments);
    setAttachmentFile(file);

    //Clear Input field
    inputRef.current.value = "";
  };

  const _getFileAsBuffer = (inputFile: any) => {
    const temporaryFileReader = new FileReader();
    return new Promise((resolve, reject) => {
      temporaryFileReader.onerror = () => {
        temporaryFileReader.abort();
        reject(new DOMException("Problem parsing input file."));
      };
      temporaryFileReader.onload = () => {
        resolve(temporaryFileReader.result);
      };
      temporaryFileReader.readAsArrayBuffer(inputFile);
    });
  };

  const _uploadFile = async () => {
    if (attachments && attachments.length > 0) {
      try {
        setIsSubmitting(true);
        // get the selected file from the input
        const file = attachmentFile; //event.target.files[0];
        console.log("file clien== ", file);

        // const path = "https://localhost:4000/sprsfa/getTkn";
        const path = process.env.SERVER_SOURCE + "/sprsfa/getTkn";
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

        // console.log("webInfo", rsData.data);
        if (!rsData.success) {
          setOpen(true);
          setSeverity("error");
          setSuccessOrErrorMessage("Api key mismatch");
          setIsSubmitting(false);
          return;
        }

        const baseURL = "https://graph.microsoft.com/v1.0";

        let fileExt = "";
        const splitted = file.name.split(".");
        if (splitted.length > 1) {
          fileExt = "." + splitted.pop();
        }
        const fileName = textInput ? textInput + fileExt : file.name; //file.name;

        // const rootUrl =
        //   "/sites/7f9b6e39-ef5e-46f8-bac1-b09c9b5a2703/drives/b!OW6bf17v-Ea6wbCcm1onA45SWzuNT4VDlSq7dVb3pCKTi3G_nWO5TI8YRD8b3e3n/items/015F3YMPSIA5JWA3Y4BFBLXMXYIDMXKEMU:/f. Email Communication/" +
        //   fileName +
        //   ":/content";

        let pathUrl = "";

        if (!subSubFolderName) {
          // pathUrl = `/sites/7f9b6e39-ef5e-46f8-bac1-b09c9b5a2703/drives/b!OW6bf17v-Ea6wbCcm1onA45SWzuNT4VDlSq7dVb3pCKTi3G_nWO5TI8YRD8b3e3n/items/${folderId}:/${subFolderName}/${fileName}:/content`;
          pathUrl = `/sites/e188129e-5eb6-4632-ac09-6f3798d411a9/drives/b!nhKI4bZeMkasCW83mNQRqbc7xjFNbW1Ck86ID6D2PaDVXFawhv_ORox8N2k5jk1J/items/${folderId}:/${subFolderName}/${fileName}:/content`;
        } else {
          // pathUrl = `/sites/7f9b6e39-ef5e-46f8-bac1-b09c9b5a2703/drives/b!OW6bf17v-Ea6wbCcm1onA45SWzuNT4VDlSq7dVb3pCKTi3G_nWO5TI8YRD8b3e3n/items/${folderId}:/${subFolderName}/${subSubFolderName}/${fileName}:/content`;
          pathUrl = `/sites/e188129e-5eb6-4632-ac09-6f3798d411a9/drives/b!nhKI4bZeMkasCW83mNQRqbc7xjFNbW1Ck86ID6D2PaDVXFawhv_ORox8N2k5jk1J/items/${folderId}:/${subFolderName}/${subSubFolderName}/${fileName}:/content`;
        }
        console.log("path====", baseURL + pathUrl);
        const pathencodedURI = encodeURI(baseURL + pathUrl);
        console.log("path===", pathencodedURI);
        const res2 = await fetch(pathencodedURI, {
          method: "PUT",
          headers: {
            Authorization: "Bearer " + rsData.data,
            "Content-Type": file.type,
          },
          body: file,
        });
        console.log("resr", res2);
        console.log(JSON.stringify(res2));
        setIsSubmitting(false);
        setSeverity("success");
        setSuccessOrErrorMessage("File upload successful!");
        setOpen(true);
        setAttachments([]);
        setTextInput("");
      } catch (error) {
        console.log("error", error);
      }
    }
  };

  const _clearItem = () => {
    setAttachments([]);
  };

  const _deleteFileItem = (index: number) => {
    let idx = index;
    // let attachments = new Array().concat(attachments);
    // attachments.splice(index, 1);
    // setAttachments(attachments);
    setAttachments([]);
  };

  const handleClose = (event?: React.SyntheticEvent | Event, reason?: string) => {
    let evt = event;
    if (reason === "clickaway") {
      return;
    }

    setOpen(false);
  };

  const handleAutoCompleteSelectionChangeClientFiles = async (event: any, newValue: any | null) => {
    console.log("clicked", newValue);
    if (!newValue) {
      setFolderId("");
      setValueClientFile(null);
      setSubFolderName("");
      setSubSubFolderName("");
      setOptionsClientFilesSubFolder([]);
      setOptionsClientFilesSubSubFolder([]);
      return;
    }
    setValueClientFile(newValue);
    let evt = event;
    try {
      console.log("Folder id : ", newValue.id);
      setFolderId(newValue.id);
      setSubFolderName("");
      setSubSubFolderName("");
      setOptionsClientFilesSubSubFolder([]);
      fetchDataClientFilesSubFolder(newValue.id);
    } catch (error) {
      console.log("move error", error);
    }
  };

  const handleAutoCompleteSelectionChangeClientFilesSubFolder = async (event: any, newValue: any | null) => {
    console.log("clicked", newValue);
    if (!newValue) {
      setSubFolderName("");
      setSubSubFolderName("");
      setOptionsClientFilesSubSubFolder([]);
      return;
    }
    let evt = event;
    try {
      console.log("Sub folder name : ", newValue.label);
      setSubFolderName(newValue.label);
      fetchDataClientFilesSubSubFolder(newValue.label);
    } catch (error) {
      console.log("Move error", error);
    }
  };

  const handleAutoCompleteSelectionChangeClientFilesSubSubFolder = async (event: any, newValue: any | null) => {
    console.log("clicked", newValue);
    if (!newValue) {
      setSubSubFolderName("");
      return;
    }
    let evt = event;
    try {
      console.log("Sub sub folder name : ", newValue.label);
      setSubSubFolderName(newValue.label);
    } catch (error) {
      console.log("Move error", error);
    }
  };

  const handleFileUpload = (file: any) => {
    console.log("file data ===", file);
    _handleFileInputChangeDropZone(file);
  };

  const handleTextInputChange = (event) => {
    setTextInput(event.target.value);
  };

  const openClientFolder = () => {
    window.location.href =
      "https://rodschubertfinancialadvice.sharepoint.com/sites/RSFATeam/Client%20files/" + valueClientFile.label;
  };

  return (
    <div className={styles.root}>
      <React.Fragment>
        <Separator />
        <Typography variant="h6" component="h6" sx={{ marginBottom: "10px" }}>
          Attachment Upload
        </Typography>

        {optionsClientFiles && optionsClientFiles.length > 0 && (
          <Autocomplete
            disablePortal
            value={valueClientFile}
            onChange={(event: any, newValue: any | null) => {
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

        {isConnectingClientFiles === true && isConnectedClientFiles === false && <Spinner size={SpinnerSize.medium} />}
        <br />
        {isConnectingClientFilesSubFolder === false &&
          optionsClientFilesSubFolder &&
          optionsClientFilesSubFolder.length > 0 && (
            <Autocomplete
              disablePortal
              value={valueClientFileSubFolder}
              onChange={(event: any, newValue: any | null) => {
                handleAutoCompleteSelectionChangeClientFilesSubFolder(event, newValue);
              }}
              inputValue={inputValueClientFileSubFolder}
              onInputChange={(event, newInputValue) => {
                const ev = event;
                setInputValueClientFileSubFolder(newInputValue);
              }}
              id="combo-box-demo-client-file-sub-folder"
              options={optionsClientFilesSubFolder}
              sx={{ width: "100%" }}
              renderInput={(params) => <TextField {...params} label="Sub Folder" />}
            />
          )}
        {isConnectingClientFilesSubFolder === true && isConnectedClientFilesSubFolder === false && (
          <Spinner size={SpinnerSize.medium} />
        )}

        <br />
        {isConnectingClientFilesSubSubFolder === false &&
          optionsClientFilesSubSubFolder &&
          optionsClientFilesSubSubFolder.length > 0 && (
            <Autocomplete
              disablePortal
              value={valueClientFileSubSubFolder}
              onChange={(event: any, newValue: any | null) => {
                handleAutoCompleteSelectionChangeClientFilesSubSubFolder(event, newValue);
              }}
              inputValue={inputValueClientFileSubSubFolder}
              onInputChange={(event, newInputValue) => {
                const ev = event;
                setInputValueClientFileSubSubFolder(newInputValue);
              }}
              id="combo-box-demo-client-file-sub-sub-folder"
              options={optionsClientFilesSubSubFolder}
              sx={{ width: "100%" }}
              renderInput={(params) => <TextField {...params} label="Sub Sub Folder" />}
            />
          )}
        {isConnectingClientFilesSubSubFolder === true && isConnectedClientFilesSubSubFolder === false && (
          <Spinner size={SpinnerSize.medium} />
        )}

        <br />
        {_renderFilesList()}
        {folderId && subFolderName && attachments.length !== 0 && (
          <>
            <TextField
              id="outlined-basic"
              label="Rename File"
              variant="outlined"
              autoComplete="off"
              value={textInput}
              onChange={handleTextInputChange}
              sx={{ width: "100%" }}
            />
          </>
        )}
        {folderId && subFolderName && attachments && attachments.length === 0 && (
          // <Link onClick={_handleFileSelect}>Add Attachment</Link>
          <DragDropFileUpload onFileUpload={handleFileUpload} />
        )}
        <input id="inputFile1" ref={inputRef} type="file" onChange={_handleFileInputChange} hidden />
      </React.Fragment>
      {folderId && subFolderName && (
        <Stack enableScopedSelectors horizontal styles={stackStyles} tokens={stackTokens}>
          <Stack.Item grow={3} styles={stackItemStyles}>
            <PrimaryButton
              text="Submit"
              onClick={_uploadFile}
              style={{ width: "100%" }}
              //allowDisabledFocus
              disabled={isSubmitting || attachments.length === 0}
              // checked={checked}
            />
            {/* <Spinner size={SpinnerSize.large} /> */}
            {isSubmitting && <Spinner size={SpinnerSize.medium} style={{ position: "absolute", left: "20%" }} />}
          </Stack.Item>
          {/* <Stack.Item grow={2} styles={stackItemStyles}>
          <DefaultButton text="Cancel" onClick={() => _clearItem()} style={{ width: "100%" }} disabled={isSubmitting} />
        </Stack.Item> */}
        </Stack>
      )}
      <Box sx={{width:"100%", textAlign: "center", padding: "0 0 30px 0"}}>
        <Button
          variant="outlined"
          disabled={!valueClientFile ? true : false}
          onClick={() => openClientFolder()}
          style={{ margin: "auto" }}
        >
          Open Client Folder
        </Button>
      </Box>
      <Snackbar open={open} autoHideDuration={6000} onClose={handleClose}>
        <Alert onClose={handleClose} severity={severity === "success" ? "success" : "error"} sx={{ width: "100%" }}>
          {successOrErrorMessage}
        </Alert>
      </Snackbar>
    </div>
  );
};

export default FileUploadToSPDocLib;
