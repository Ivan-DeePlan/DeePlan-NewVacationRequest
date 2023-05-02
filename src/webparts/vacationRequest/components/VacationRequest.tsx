import * as React from "react";
import styles from "./VacationRequest.module.scss";
import "./VacationRequest.css";
import getSP from "../PnPjsConfig";
import { IVacationRequestProps } from "./IVacationRequestProps";
import { IVacationRequestStates } from "./IVacationRequestStates";
import { Constants } from "../Models/Constants";
import { AiOutlineSend } from "react-icons/ai";
import { IoCloseCircleOutline } from "react-icons/io5";
import {
  MuiPickersUtilsProvider,
  KeyboardDatePicker,
} from "@material-ui/pickers";
import DateFnsUtils from "@date-io/date-fns";
import { he } from "date-fns/locale";
import { CacheProvider } from "@emotion/react";
import createCache from "@emotion/cache";
import stylisRTLPlugin from "stylis-plugin-rtl";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TextField } from "@material-ui/core";
import Checkbox from "@mui/material/Checkbox";
import FormControlLabel from "@mui/material/FormControlLabel";
import {
  Row,
  Col,
  Container,
  Alert,
  Popover,
  PopoverHeader,
  PopoverBody,
} from "reactstrap";
import Box from "@mui/material/Box";
import { Button, TextField as MaterialTextField } from "@mui/material";
import { Autocomplete } from "@mui/material";

export default class VacationRequest extends React.Component<
  IVacationRequestProps,
  IVacationRequestStates
> {
  sp = getSP(this.props.context);

  //* use for direction rtl
  cacheRtl = createCache({
    key: "muirtl",
    stylisPlugins: [stylisRTLPlugin],
  });

  constructor(props: IVacationRequestProps) {
    super(props);
    this.state = {
      IsLoading: false,
      popoverOpen: false,
      isFutuerData: false,
      siteGroups: "",
      isPolicyVacationChacke: false,
      isBiggerThenStartData: false,
      CompanyDepartments: [],
      isManager: "No",
      approvalData: {
        approvalDManagerStatus: "בחר",
        approvalManagerStatus: "בחר",
      },
      requestValues: {
        currentUser: null,
        currentGroup: "",
        Url: "https://projects1.sharepoint.com/sites/DeeplanPortal/SitePages/EditVacationRequest.aspx?FormID=",
        UrlDescription: "",
        RequestDate: new Date(),
        CompanyDepartmenId: 0,
        CompanyManagerId: 0,
        DepartmentManagerId: 0,
        FromDate: null,
        ToDate: null,
        numberOfDays: 0,
        policyVacationChacke: false,
        haveVacationDays: "",
        vacationDuringActiveProject: "",
        vacationDuringActiveProjectEX: "",
      },
    };
  }

  //* PeoplePicker styles
  public _styles: any = {
    root: { maxWidth: "300px" },
    input: { with: "100%" },
  };

  componentDidMount() {
    //* Start Loader
    this.setState({
      IsLoading: true,
    });

    //* Reset all the values in the form
    this.ResetForm();
  }

  //*Change the current state of the requestValue <= CompanyDepartmen and departmentManagerEmail
  CompanyDepartmentHandler = (Department: string) => {
    const department = this.state.CompanyDepartments.find(
      (company: any) => company.Title === Department
    );

    this.setState({
      requestValues: {
        ...this.state.requestValues,
        CompanyDepartmenId: department.Id,
        DepartmentManagerId: department.DepartmentManagerId,
        currentGroup: Department,
      },
    });

    if (
      this.state.requestValues.currentUser.Id === department.DepartmentManagerId
    ) {
      this.setState({
        isManager: "Yes",
      });
    } else {
      this.setState({
        isManager: "No",
      });
    }
  };

  ResetForm = async () => {
    await this.sp.web.lists
      .getByTitle("VacationRequests")
      .items()
      .catch((error) => {
        console.log(error);
      })
      .then((object: any) => {
        console.log(object);
      });
    //* get all the items from a list Company Departments
    await this.sp.web.lists
      .getByTitle("CompanyDepartments")
      .items()
      .catch((error) => {
        console.log(error);
      })
      .then((Company: any) => {
        this.setState({
          CompanyDepartments: Company,
        });
        const department = this.state.CompanyDepartments.find(
          (company: any) => company.Title === 'מנכ"ל'
        );

        this.setState({
          requestValues: {
            ...this.state.requestValues,
            CompanyManagerId: department.DepartmentManagerId,
          },
        });
      });

    //* get current user and department
    await this.sp.web.currentUser
      .select()()
      .catch((error) => {
        console.log(error);
      })
      .then(async (user: any) => {
        this.setState({
          requestValues: {
            ...this.state.requestValues,
            currentUser: user,
          },
        });

        this.sp.web.siteUsers
          .getById(user.Id)
          .groups()
          .then((groups: any) => {
            if (groups.some((group: any) => group.Title === "מחלקת פיתוח")) {
              this.CompanyDepartmentHandler("פיתוח");
            } else if (
              groups.some((group: any) => group.Title === "מחלקת גיוס")
            ) {
              this.CompanyDepartmentHandler("גיוס");
            } else if (
              groups.some((group: any) => group.Title === "מחלקת חומרה")
            ) {
              this.CompanyDepartmentHandler("חומרה");
            } else {
              this.CompanyDepartmentHandler("אחר");
            }
          });
      });

    this.setState({
      IsLoading: false,
    });
  };

  //* if haveVacationDays = "לא יודע" need to check else set all values and go to the list
  onSubmithandler = (event: any) => {
    event.preventDefault();
    if (
      this.state.requestValues.haveVacationDays === "לא יודע" &&
      !this.state.requestValues.policyVacationChacke
    ) {
      this.setState({ isPolicyVacationChacke: true });
    } else if (this.state.requestValues) {
      try {
        this.sp.web.lists
          .getByTitle("VacationRequests")
          .items.add({
            RequestedByUserId: this.state.requestValues.currentUser.Id,
            VecationRequestDate: this.state.requestValues.RequestDate,
            DepartmentId: this.state.requestValues.CompanyDepartmenId,
            DManager: this.state.approvalData.approvalDManagerStatus,
            CManager: this.state.approvalData.approvalManagerStatus,
            IsManager: this.state.isManager,
            ApprovalsId:
              this.state.requestValues.DepartmentManagerId !== null
                ? this.state.requestValues.DepartmentManagerId
                : null,
            CompanyManagerId: this.state.requestValues.CompanyManagerId,
            FromDate: this.state.requestValues.FromDate,
            ToDate: this.state.requestValues.ToDate,
            NumberOfVacationDays: this.state.requestValues.numberOfDays,
            haveVacationDaysLeft: this.state.requestValues.haveVacationDays,
            VacationDuringActiveProject:
              this.state.requestValues.vacationDuringActiveProject,
            vacationDuringActiveProjectEX:
              this.state.requestValues.vacationDuringActiveProjectEX,
          })
          .catch((error) => console.log(error))
          .then((listItem: any) => {
            //* get the created list item Id to set the hyperLink for editing
            this.sp.web.lists
              .getByTitle("VacationRequests")
              .items.getById(listItem.data.Id)
              .update({
                appForm: {
                  Description:
                    this.state.requestValues.currentUser.Title +
                    " " +
                    new Date().toLocaleDateString("pt-PT"),
                  Url: this.state.requestValues.Url + listItem.data.Id,
                },
              });
            window.location.href = Constants.ListView;
          });
      } catch (error) {
        console.log(error);
      }
    }
  };

  //* go back to the list
  onCanclehandler = async () => {
    window.location.href = Constants.ListView;
  };

  //* sweet alert
  PopOverToggle = () => {
    this.setState({
      popoverOpen: !this.state.popoverOpen,
    });
  };

  public render(): React.ReactElement<IVacationRequestProps> {
    // //*Change the current state of the requestValue <= currentUser
    const userHandler = async (e: any) => {
      const username = e[0].secondaryText;
      await this.sp.web.ensureUser(username).then((user: any) => {
        this.setState({
          requestValues: {
            ...this.state.requestValues,
            currentUser: user.data,
          },
        });
        if (user.data.Id === this.state.requestValues.DepartmentManagerId) {
          this.setState({
            isManager: "Yes",
          });
        } else {
          this.setState({
            isManager: "No",
          });
        }
        this.sp.web.siteUsers
          .getById(user.data.Id)
          .groups()
          .then((groups: any) => {
            if (groups.some((group: any) => group.Title === "מחלקת פיתוח")) {
              this.CompanyDepartmentHandler("פיתוח");
            } else if (
              groups.some((group: any) => group.Title === "מחלקת גיוס")
            ) {
              this.CompanyDepartmentHandler("גיוס");
            } else if (
              groups.some((group: any) => group.Title === "מחלקת חומרה")
            ) {
              this.CompanyDepartmentHandler("חומרה");
            } else {
              this.CompanyDepartmentHandler("אחר");
            }
          });
      });
    };

    // //*Change the current Department
    const DepartmentHandler = (e: any) => {
      this.CompanyDepartmentHandler(e.target.innerText);
    };

    // //*Change the current state of the requestValue <= FromDate
    const FromDateHandler = (e: any) => {
      const today = new Date();
      if (e < today.getTime()) {
        this.setState({ isFutuerData: true });
      } else
        this.setState({
          isFutuerData: false,
          requestValues: {
            ...this.state.requestValues,
            FromDate: e,
          },
        });
    };
    // //*Change the current state of the requestValue <= ToDate
    const ToDateHandler = (e: any) => {
      if (e <= this.state.requestValues.FromDate) {
        this.setState({ isBiggerThenStartData: true });
      } else {
        this.setState({
          isBiggerThenStartData: false,
          requestValues: {
            ...this.state.requestValues,
            ToDate: e,
          },
        });
      }
    };
    // //*Change the current state of the requestValue <= numberOfDays
    const numberOfDaysHandler = (e: any) => {
      this.setState({
        requestValues: {
          ...this.state.requestValues,
          numberOfDays: e.target.value,
        },
      });
    };
    // //*Change the current state of the requestValue <= haveVacationDays
    const haveVacationDaysHandler = (e: any) => {
      this.setState({
        requestValues: {
          ...this.state.requestValues,
          haveVacationDays: e.target.innerText,
        },
      });
    };
    // //*Change the current state of the request Value <= policyVacationChacke
    const policyVacationChackeChange = (e: any) => {
      this.setState({
        requestValues: {
          ...this.state.requestValues,
          policyVacationChacke: e.target.checked,
        },
      });
    };
    // //*Change the current state of the requestValue <= vacationDuringActiveProject
    const vacationDuringActiveProjectHandler = (e: any) => {
      this.setState({
        requestValues: {
          ...this.state.requestValues,
          vacationDuringActiveProject: e.target.innerText,
        },
      });
    };
    // //*Change the current state of the requestValue <= vacationDuringActiveProjectEX
    const vacationDuringActiveProjectEX = (e: any) => {
      this.setState({
        requestValues: {
          ...this.state.requestValues,
          vacationDuringActiveProjectEX: e.target.value,
        },
      });
    };
    return (
      <div className={styles.NewRequest} dir="rtl">
        <CacheProvider value={this.cacheRtl}>
          <div className="EONewFormContainer">
            <div className="EOHeader">
              <div className="EOHeaderContainer">
                <span className="EOHeaderText">בקשה לחופשה</span>
              </div>
              <div className="EOLogoContainer"></div>
            </div>
            {this.state.IsLoading && (
              <div className="SpinnerComp">
                <div className="loading-screen">
                  <div className="loader-wrap">
                    <span className="loader-animation"></span>
                    <div className="loading-text">
                      <span className="letter">ב</span>
                      <span className="letter">ט</span>
                      <span className="letter">ע</span>
                      <span className="letter">י</span>
                      <span className="letter">נ</span>
                      <span className="letter">ה</span>
                    </div>
                  </div>
                </div>
              </div>
            )}
            <form onSubmit={this.onSubmithandler}>
              <Row className="mt-4">
                {" "}
                <Col xs={3}>
                  <p>
                    <span className="required-dot">*</span>שם העובד:
                  </p>
                </Col>
                <Col xs={9}>
                  <PeoplePicker
                    context={this.props.context}
                    personSelectionLimit={1}
                    showtooltip={true}
                    required={true}
                    styles={this._styles}
                    defaultSelectedUsers={[
                      this.state.requestValues.currentUser &&
                        this.state.requestValues.currentUser.Email,
                    ]}
                    onChange={userHandler}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                  />
                </Col>
              </Row>
              <Row className="mt-4">
                <Col xs={3}>
                  {" "}
                  <p>תאריך הבקשה:</p>
                </Col>
                <Col xs={9}>
                  {" "}
                  <MuiPickersUtilsProvider utils={DateFnsUtils} locale={he}>
                    <KeyboardDatePicker
                      variant="inline"
                      format="dd/MM/yyyy"
                      id="Date Picker"
                      onChange={null}
                      readOnly
                      disabled
                      value={this.state.requestValues.RequestDate}
                      style={{
                        width: 300,
                        marginLeft: 5,
                      }}
                      KeyboardButtonProps={{
                        "aria-label": "change date",
                      }}
                      InputProps={{
                        style: {
                          fontSize: 16,
                          height: 30,
                        },
                      }}
                    />
                  </MuiPickersUtilsProvider>
                </Col>
              </Row>
              <Row className="mt-4">
                <Col xs={3}>
                  {" "}
                  <p>מחלקה:</p>
                </Col>
                <Col xs={9}>
                  {" "}
                  <div className="box">
                    <Autocomplete
                      id="CompanyDepartments"
                      sx={{ width: "300px" }}
                      style={{ direction: "rtl" }}
                      onChange={DepartmentHandler}
                      value={this.state.requestValues.currentGroup}
                      disabled
                      options={
                        this.state.CompanyDepartments &&
                        this.state.CompanyDepartments.filter(
                          (Department: any) => Department.Title !== 'מנכ"ל'
                        ).map((Department: any) => Department.Title)
                      }
                      renderInput={(params) => (
                        <TextField
                          required
                          style={{ padding: 0, textAlign: "center" }}
                          variant="outlined"
                          {...params}
                        />
                      )}
                    />
                  </div>
                </Col>
              </Row>
              <Row className="mt-4">
                <Col xs={3}>
                  <p>
                    <span className="required-dot">*</span>תאריכים:
                  </p>
                </Col>
                <Col xs={3}>
                  <MuiPickersUtilsProvider utils={DateFnsUtils} locale={he}>
                    <KeyboardDatePicker
                      variant="inline"
                      format="dd/MM/yyyy"
                      id="Date Picker"
                      placeholder="מתאריך"
                      required
                      onChange={FromDateHandler}
                      value={this.state.requestValues.FromDate}
                      autoOk={true}
                      style={{
                        width: 150,
                        marginLeft: 5,
                      }}
                      KeyboardButtonProps={{
                        "aria-label": "change date",
                      }}
                      InputProps={{
                        style: {
                          fontSize: 16,
                          height: 30,
                          width: "200px",
                        },
                      }}
                    />
                  </MuiPickersUtilsProvider>
                </Col>
                <Col xs={3}>
                  <MuiPickersUtilsProvider utils={DateFnsUtils} locale={he}>
                    <KeyboardDatePicker
                      variant="inline"
                      format="dd/MM/yyyy"
                      id="FromDate"
                      placeholder="עד תאריך"
                      required
                      onChange={ToDateHandler}
                      value={this.state.requestValues.ToDate}
                      autoOk={true}
                      style={{
                        width: 150,
                        marginLeft: 5,
                      }}
                      KeyboardButtonProps={{
                        "aria-label": "change date",
                      }}
                      InputProps={{
                        style: {
                          fontSize: 16,
                          height: 30,
                          width: "200px",
                        },
                      }}
                    />
                  </MuiPickersUtilsProvider>
                </Col>
              </Row>
              <Row className="mt-3 justify-content-center">
                {this.state.isFutuerData && (
                  <Col xs={5}>
                    <Alert className="text-center" color="danger">
                      נא לבחור תאריך עתידי
                    </Alert>
                  </Col>
                )}
              </Row>
              <Row className="mt-3 justify-content-center">
                {this.state.isBiggerThenStartData && (
                  <Col xs={5}>
                    <Alert className="text-center" color="danger">
                      התאריך חייב להיות גדול ממתאריך
                    </Alert>
                  </Col>
                )}
              </Row>
              <Row className="mt-4">
                <Col xs={3}>
                  {" "}
                  <p>
                    <span className="required-dot">*</span>מספר ימי חופשה בפועל:
                  </p>
                </Col>
                <Col xs={9}>
                  {" "}
                  <TextField
                    required
                    style={{
                      width: 300,
                    }}
                    InputProps={{ inputProps: { min: 0 } }}
                    type="number"
                    onChange={numberOfDaysHandler}
                  ></TextField>
                </Col>
              </Row>
              <Row className="mt-4">
                <Col xs={3}>
                  {" "}
                  <p>
                    <span className="required-dot">*</span>האם נותרו לי ימי
                    חופשה?
                  </p>
                </Col>
                <Col xs={9}>
                  {" "}
                  <div className="box">
                    <Autocomplete
                      id="haveVacationDays"
                      sx={{ width: "300px" }}
                      onChange={haveVacationDaysHandler}
                      options={["כן", "לא", "לא יודע"]}
                      renderInput={(params) => (
                        <TextField
                          required
                          style={{ padding: 0, textAlign: "center" }}
                          variant="outlined"
                          {...params}
                        />
                      )}
                    />
                  </div>
                </Col>
              </Row>{" "}
              {this.state.requestValues.haveVacationDays === "לא יודע" && (
                <Row className="mt-3 justify-content-start">
                  <Col xs={3}></Col>
                  <Col xs={9}>
                    <FormControlLabel
                      value={this.state.isPolicyVacationChacke}
                      control={<Checkbox />}
                      label="אני מאשר/ת בזה לחברה שאם אין לי די ימי חופשה צבורה החברה
                      תוכל לקזז את החופשה שתאושר משכרי"
                      onChange={policyVacationChackeChange}
                      labelPlacement="end"
                    />
                  </Col>
                </Row>
              )}
              <Row className="mt-3 justify-content-center">
                {this.state.isPolicyVacationChacke && (
                  <Col xs={5}>
                    <Alert className="text-center" color="danger">
                      נא לאשר את תנאי הבקשה
                    </Alert>
                  </Col>
                )}
              </Row>
              {this.state.requestValues.CompanyDepartmenId === 1 && (
                <Row className="mt-3">
                  <Col xs={3}>
                    <p>האם היציאה לחופשה מתקיימת בזמן פרויקט פעיל?</p>
                  </Col>
                  <Col xs={9}>
                    {" "}
                    <div className="box">
                      <Autocomplete
                        className="mt-2"
                        id="haveVacationDays"
                        sx={{ width: "300px" }}
                        onChange={vacationDuringActiveProjectHandler}
                        options={["כן", "לא", "לא יודע", "אחר"]}
                        renderInput={(params) => (
                          <TextField
                            style={{ padding: 0, textAlign: "center" }}
                            variant="outlined"
                            {...params}
                          />
                        )}
                      />
                    </div>
                  </Col>
                </Row>
              )}
              {this.state.requestValues.vacationDuringActiveProject === "אחר" &&
                this.state.requestValues.CompanyDepartmenId === 1 && (
                  <Row>
                    <Col xs={3}></Col>
                    <Col xs={9} style={{ padding: "0px" }}>
                      <Box
                        component="form"
                        sx={{
                          "& .MuiTextField-root": { m: 1, width: "70ch" },
                        }}
                        noValidate
                        autoComplete="off"
                      >
                        <div>
                          <MaterialTextField
                            dir="rtl"
                            id="outlined-multiline-static"
                            onChange={vacationDuringActiveProjectEX}
                            label="פרט"
                            multiline
                            rows={4}
                          />
                        </div>
                      </Box>
                    </Col>
                  </Row>
                )}
              <Container className="mt-5 mb-5">
                <Row className="justify-content-md-center">
                  <Col xs={1} className="text-center">
                    <Button
                      variant="contained"
                      id="Popover1"
                      color="error"
                      onClick={this.PopOverToggle}
                      endIcon={<IoCloseCircleOutline />}
                    >
                      ביטול
                    </Button>
                    <Popover
                      flip
                      placement="top"
                      target="Popover1"
                      toggle={this.PopOverToggle}
                      isOpen={this.state.popoverOpen}
                    >
                      <PopoverHeader className="text-center">
                        ?האם אתה בטוח
                      </PopoverHeader>
                      <PopoverBody>
                        <div>
                          {" "}
                          <Button
                            variant="contained"
                            style={{ backgroundColor: "#84C792" }}
                            onClick={this.onCanclehandler}
                          >
                            כן
                          </Button>
                          &nbsp;&nbsp;
                          <Button
                            variant="contained"
                            color="error"
                            onClick={this.PopOverToggle}
                          >
                            לא
                          </Button>
                        </div>
                      </PopoverBody>
                    </Popover>
                  </Col>
                  <Col xs={1} className="text-center">
                    <Button
                      variant="contained"
                      style={{ backgroundColor: "#84C792" }}
                      type="submit"
                      endIcon={<AiOutlineSend className="RotatedIcon" />}
                    >
                      שמירה
                    </Button>
                  </Col>
                </Row>
              </Container>
            </form>
          </div>
        </CacheProvider>
      </div>
    );
  }
}
