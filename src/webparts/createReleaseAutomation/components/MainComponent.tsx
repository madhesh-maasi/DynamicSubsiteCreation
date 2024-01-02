import * as React from "react";
import { sp, Web } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import {
  DefaultButton,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Text,
  TextField,
} from "@fluentui/react";
import "@pnp/sp/navigation";
import Loader from "./Loader";
import styles from "./CreateReleaseAutomation.module.scss";
const MainComponent = (props) => {
  const [value, setValue] = React.useState("");
  const [isLoader, setIsLoader] = React.useState<boolean>(false);
  const [Data, setData] = React.useState([]);
  const [masterData, setMasterData] = React.useState([]);
  const [error, setError] = React.useState("");
  const [reRender, setReRedender] = React.useState(false);
  //   const subsiteCreationInfo: WebCreationInformation = {
  //     Title: "Subsite Title",
  //     Url: "subsite-url",
  //     WebTemplate: "STS#0", // Use the appropriate template
  //   };
  const pageurl = props.context.pageContext.web.absoluteUrl;
  const sitePages = [
    {
      Title: "Release Information",
      PageType: "Article",
      PageTitle: "Release Information",
    },
    { Title: "Release Notes", PageType: "Article", PageTitle: "Release Notes" },
    {
      Title: "Features Mapping",
      PageType: "Article",
      PageTitle: "Features Mapping",
    },
    {
      Title: "Supported Platforms",
      PageType: "Article",
      PageTitle: "Supported Platforms",
    },
    {
      Title: "Arch & Design ",
      PageType: "Article",
      PageTitle: "Arch & Design",
    },
    {
      Title: "Cloud Architecture",
      PageType: "Article",
      PageTitle: "Cloud Architecture",
    },
    {
      Title: "Overall Architecture",
      PageType: "Article",
      PageTitle: "Overall Architecture",
    },
    {
      Title: "High Availability",
      PageType: "Article",
      PageTitle: "High Availability",
    },
    {
      Title: "Alerting and Monitoring",
      PageType: "Article",
      PageTitle: "Alerting and Monitoring",
    },
    { Title: "Logging ", PageType: "Article", PageTitle: "Logging" },
    {
      Title: "On Premise Architecture",
      PageType: "Article",
      PageTitle: "On Premise Architecture",
    },

    {
      Title: "Overall Architecture",
      PageType: "Article",
      PageTitle: "Overall Architecture",
    },
    {
      Title: "High Availability",
      PageType: "Article",
      PageTitle: "High Availability",
    },
    {
      Title: "Alerting and Monitoring",
      PageType: "Article",
      PageTitle: "Alerting and Monitoring",
    },
    { Title: "Logging ", PageType: "Article", PageTitle: "Logging" },
    {
      Title: "CTI Integrations",
      PageType: "Article",
      PageTitle: "CTI Integrations",
    },
    {
      Title: "Error Codes / Disposition Codes",
      PageType: "Article",
      PageTitle: "Error Codes / Disposition Codes",
    },
    {
      Title: "Database Design",
      PageType: "Article",
      PageTitle: "Database Design",
    },
    {
      Title: " API Sawgger Documentation",
      PageType: "Article",
      PageTitle: "API Sawgger Documentation",
    },
    {
      Title: "Session API Documentation",
      PageType: "Article",
      PageTitle: "Session API Documentation",
    },
    {
      Title: "Receive Communication Documentation",
      PageType: "Article",
      PageTitle: "Receive Communication Documentation",
    },
    { Title: "Deployment", PageType: "Article", PageTitle: "Deployment" },
    {
      Title: "Cloud Deployment",
      PageType: "Article",
      PageTitle: "Cloud Deployment",
    },
    {
      Title: "CDK Scripts",
      PageType: "Article",
      PageTitle: "CDK Deployment in AWS",
    },
    {
      Title: "PM2 Install",
      PageType: "Article",
      PageTitle: "PM2 Install",
    },
    {
      Title: "On Premise Deployment",
      PageType: "Article",
      PageTitle: "On Premise Deployment",
    },
    {
      Title: "Redhat Install",
      PageType: "Article",
      PageTitle: "Redhat Install",
    },
    {
      Title: "Ubuntu Install",
      PageType: "Article",
      PageTitle: "Ubuntu Install",
    },
    {
      Title: "Windows Install",
      PageType: "Article",
      PageTitle: " Windows Install",
    },
    {
      Title: "Cisco Finesse Install",
      PageType: "Article",
      PageTitle: "Cisco Finesse Install",
    },
    {
      Title: "Configuration",
      PageType: "Article",
      PageTitle: "Configuration",
    },
    {
      Title: "Mirth Configuration",
      PageType: "Article",
      PageTitle: "Mirth Configuration",
    },
    {
      Title: "Settings File",
      PageType: "Article",
      PageTitle: "Settings File",
    },
    {
      Title: "Nginx Configuration",
      PageType: "Article",
      PageTitle: "Nginx Configuration",
    },
    {
      Title: "DB Configuration",
      PageType: "Article",
      PageTitle: "DB Configuration",
    },
    {
      Title: "Alerting and Monitoring Setup",
      PageType: "Article",
      PageTitle: "Alerting and Monitoring Setup",
    },
    {
      Title: "Data Migration",
      PageType: "Article",
      PageTitle: "Data Migration",
    },
    {
      Title: "Provider Data Migration",
      PageType: "Article",
      PageTitle: "Provider Data Migration",
    },
    {
      Title: "Calling Destination Migration",
      PageType: "Article",
      PageTitle: "Calling Destination Migration",
    },
    {
      Title: "Support and Troubleshooting",
      PageType: "Article",
      PageTitle: "Support and Troubleshooting",
    },
    {
      Title: "Troubleshooting Guide",
      PageType: "Article",
      PageTitle: "Troubleshooting Guide",
    },
    {
      Title: "API Mapping",
      PageType: "Article",
      PageTitle: "API Mapping",
    },
    {
      Title: "Error Codes and Disposition Codes",
      PageType: "Article",
      PageTitle: "Error Codes and Disposition Codes",
    },
    { Title: "FAQs", PageType: "Article", PageTitle: "FAQs" },
    {
      Title: "Common Epic APIs",
      PageType: "Article",
      PageTitle: "Common Epic APIs",
    },
  ];

  const SubsiteCreate = async () => {
    setData([...Data, value]);
    // Replace these values with your actual site URL and subsite details
    // let x = "https://chandrudemo.sharepoint.com/sites/CreateReleaseAutomation";
    const siteUrl = props.context.pageContext.web.absoluteUrl;
    const subsiteUrl = value;
    const subsiteTitle = value;
    const subsiteDescription = "Description for the new subsite";
    const WebTemplate = "STS#3";
    let ReleaseID: number;
    if (value.trim() != "") {
      try {
        const web = Web(siteUrl);

        // Create subsite using the REST API
        await sp.web.webs
          .add(subsiteTitle, subsiteUrl, subsiteDescription, WebTemplate)
          .then(async (res) => {
            // createSitePages();

            const xweb = props.context.pageContext.web.absoluteUrl;
            // console.log(xweb, "siteurl");
            const xxweb = Web(xweb);
            const result = await xxweb.addClientsidePage(value, "Article");
            web.navigation.quicklaunch.get().then((res) => {
              console.log(res);
              ReleaseID = res.filter((li) => li.Title == "Releases")[0].Id;
            });
            await result
              .save()
              .then(async (res) => {
                const xweb1 = props.context.pageContext.web.absoluteUrl;
                const subpage =
                  props.context.pageContext.web.absoluteUrl + "/" + value;
                const xxweb = Web(xweb1);
                await web.navigation.quicklaunch
                  .getById(ReleaseID)
                  .children.add(value, subpage, true)
                  .then((res) => {
                    createSitePages();
                  })
                  .catch((err) => {
                    setIsLoader(false);
                  });
              })
              .catch((err) => {
                setIsLoader(false);
              });
          })
          .catch((err) => {
            setIsLoader(false);
          });

        // console.log("Subsite created successfully.");
      } catch (error) {
        setIsLoader(false);

        console.error("Error creating subsite:", error);
      }
    } else {
      setIsLoader(false);
      return;
    }
  };

  //site page
  const createSitePages = async () => {
    const xweb = props.context.pageContext.web.absoluteUrl + "/" + value;
    // console.log(xweb, "siteurl");
    const xxweb = Web(xweb);

    for (let i: number = 0; sitePages.length > i; i++) {
      const result = await xxweb.addClientsidePage(
        sitePages[i].Title,
        sitePages[i].PageTitle,
        "Article"
      );

      await result
        .save()
        .then((_res) => {
          // console.log("siteres", _res);

          if (sitePages.length === i + 1) {
            createNavigationTree();
          }
        })
        .catch((err: any) => {
          setIsLoader(false);
          console.log("err > ", err);
        });
    }
  };

  //navigation

  const navigationItems = [
    {
      title: "Release Information",
      //   url: "/sites/POCforLeftNav/Test001/SitePages/Test.aspx",
      //   url: `/sites/CreateReleaseAutomation/${value}/SitePages/Release-Information.aspx`,
      url: `${pageurl}/${value}/SitePages/Release-Information.aspx`,
      // url: "/sites/POCforLeftNav/Test001/SitePages/Test.aspx",
      isExternal: false,
      sequence: 1,
      children: [
        {
          title: "Release Notes",
          url: `${pageurl}/${value}/SitePages/Release-Notes.aspx`,
          isExternal: false,
          sequence: 1,
          children: [
            {
              title: "Features Mapping",
              url: `${pageurl}/${value}/SitePages/Features-Mapping.aspx`,
              isExternal: false,
              sequence: 1,
              children: [],
            },
          ],
        },
        {
          title: "Supported Platforms",
          url: `${pageurl}/${value}/SitePages/Supported-Platforms.aspx`,
          isExternal: false,
          sequence: 2,
          children: [],
        },
        // {
        //   title: "Test2",
        //   url: "/sites/POCforLeftNav/Test001/SitePages/Test2.aspx",
        //   isExternal: false,
        //   sequence: 2,
        //   children: [],
        // },
      ],
    },
    {
      title: "Arch & Design",
      url: `${pageurl}/${value}/SitePages/Arch-&-Design.aspx`,
      //   url: "/sites/POCforLeftNav/Test001/SitePages/Test.aspx",
      isExternal: false,
      sequence: 1,
      children: [
        {
          title: "Cloud Architecture",
          url: `${pageurl}/${value}/SitePages/Cloud-Architecture.aspx`,
          isExternal: false,
          sequence: 1,
          children: [
            {
              title: "Overall Architecture",
              url: `${pageurl}/${value}/SitePages/Overall-Architecture.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "High Availability",
              url: `${pageurl}/${value}/SitePages/High-Availability.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Alerting and Monitoring",
              url: `${pageurl}/${value}/SitePages/Alerting-and-Monitoring.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Logging",
              url: `${pageurl}/${value}/SitePages/Logging.aspx`,
              isExternal: false,
              sequence: 1,
            },
          ],
        },
        {
          title: "On Premise Architecture",
          url: `${pageurl}/${value}/SitePages/On-Premise-Architecture.aspx`,
          isExternal: false,
          sequence: 2,
          children: [
            {
              title: "Overall Architecture",
              url: `${pageurl}/${value}/SitePages/Overall-Architecture(1).aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "High Availability",
              url: `${pageurl}/${value}/SitePages/High-Availability(1).aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Alerting and Monitoring",
              url: `${pageurl}/${value}/SitePages/Alerting-and-Monitoring(1).aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Logging",
              url: `${pageurl}/${value}/SitePages/Logging(1).aspx`,
              isExternal: false,
              sequence: 1,
            },
          ],
        },
        {
          title: "CTI Integrations",
          url: `${pageurl}/${value}/SitePages/CTI-Integrations.aspx`,
          isExternal: false,
          sequence: 2,
          children: [
            {
              title: "Error Codes / Disposition Codes",
              url: `${pageurl}/${value}/SitePages/Error-codes---Disposition-codes.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Database Design",
              url: `${pageurl}/${value}/SitePages/Database-design.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "API Sawgger Documentation",
              url: `${pageurl}/${value}/SitePages/API-Sawgger-documentation.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Session API Documentation",
              url: `${pageurl}/${value}/SitePages/Session-API-documentation.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Receive Communication Documentation",
              url: `${pageurl}/${value}/SitePages/Receive-Communication-Documentation.aspx`,
              isExternal: false,
              sequence: 1,
            },
          ],
        },
      ],
    },
    {
      title: "Deployment",
      url: `${pageurl}/${value}/SitePages/Deployment.aspx`,
      //   url: "/sites/POCforLeftNav/Test001/SitePages/Test.aspx",
      isExternal: false,
      sequence: 1,
      children: [
        {
          title: "Cloud Deployment",
          url: `${pageurl}/${value}/SitePages/Cloud-Deployment.aspx`,
          isExternal: false,
          sequence: 1,
          children: [
            {
              title: "CDK Scripts",
              url: `${pageurl}/${value}/SitePages/CDK-Scripts.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "PM2 Install",
              url: `${pageurl}/${value}/SitePages/PM2-Install.aspx`,
              isExternal: false,
              sequence: 1,
            },
            // {
            //   title: "3.Alerting and Monitoring",
            //   url: "/sites/POCforLeftNav/Test001/SitePages/Alerting-and-Monitoring.aspx",
            //   isExternal: false,
            //   sequence: 1,
            // },
            // {
            //   title: "Logging",
            //   url: "/sites/POCforLeftNav/Test001/SitePages/Logging.aspx",
            //   isExternal: false,
            //   sequence: 1,
            // },
          ],
        },
        {
          title: "On Premise Deployment",
          url: `${pageurl}/${value}/SitePages/On-Premise-Deployment.aspx`,
          isExternal: false,
          sequence: 2,
          children: [
            {
              title: "Redhat Install",
              url: `${pageurl}/${value}/SitePages/Redhat-install.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Ubuntu Install",
              url: `${pageurl}/${value}/SitePages/Ubuntu-Install.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Windows Install",
              url: `${pageurl}/${value}/SitePages/Windows-Install.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Cisco Finesse Install",
              url: `${pageurl}/${value}/SitePages/Cisco-Finesse-Install.aspx`,
              isExternal: false,
              sequence: 1,
            },
          ],
        },
        {
          title: "Configuration",
          url: `${pageurl}/${value}/SitePages/Configuration.aspx`,
          isExternal: false,
          sequence: 2,
          children: [
            {
              title: "Mirth-Configuration",
              url: `${pageurl}/${value}/SitePages/Mirth-Configuration.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Settings File",
              url: `${pageurl}/${value}/SitePages/Settings-file.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Nginx Configuration",
              url: `${pageurl}/${value}/SitePages/nginx-configuration.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "DB Configuration",
              url: `${pageurl}/${value}/SitePages/DB-configuration.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Alerting and Monitoring Setup",
              url: `${pageurl}/${value}/SitePages/Alerting-and-Monitoring-setup.aspx`,
              isExternal: false,
              sequence: 1,
            },
          ],
        },
        {
          title: "Data Migration",
          url: `${pageurl}/${value}/SitePages/Data-Migration.aspx`,
          isExternal: false,
          sequence: 2,
          children: [
            {
              title: "Provider Data Migration",
              url: `${pageurl}/${value}/SitePages/Provider-data-migration.aspx`,
              isExternal: false,
              sequence: 1,
            },
            {
              title: "Calling Destination Migration",
              url: `${pageurl}/${value}/SitePages/Calling-Destination-migration.aspx`,
              isExternal: false,
              sequence: 2,
            },
          ],
        },
      ],
    },
    {
      title: "Support and Troubleshooting",
      url: `${pageurl}/${value}/SitePages/Support-and-Troubleshooting.aspx`,
      //   url: "/sites/POCforLeftNav/Test001/SitePages/Test.aspx",
      isExternal: false,
      sequence: 1,
      children: [
        {
          title: "Troubleshooting Guide",
          url: `${pageurl}/${value}/SitePages/Troubleshooting-guide.aspx`,
          isExternal: false,
          sequence: 1,
          children: [],
        },
        {
          title: "API Mapping",
          url: `${pageurl}/${value}/SitePages/API-mapping.aspx`,
          isExternal: false,
          sequence: 2,
          children: [],
        },
        {
          title: "Error Codes and Disposition Codes",
          url: `${pageurl}/${value}/SitePages/Error-Codes-and-Disposition-codes.aspx`,
          isExternal: false,
          sequence: 2,
          children: [],
        },
        {
          title: "FAQs",
          url: `${pageurl}/${value}/SitePages/FAQs.aspx`,
          isExternal: false,
          sequence: 2,
          children: [],
        },
        {
          title: " Common Epic APIs",
          url: `${pageurl}/${value}/SitePages/Common-Epic-APIs.aspx`,
          isExternal: false,
          sequence: 2,
          children: [],
        },
      ],
    },
  ];
  const createNavigationTree = async () => {
    const xweb1 = props.context.pageContext.web.absoluteUrl + "/" + value;
    // console.log(xweb1, "siteurl");
    const xxweb = Web(xweb1);
    for (let i: number = 0; navigationItems.length > i; i++) {
      //   await sp.web.navigation.quicklaunch
      await xxweb.navigation.quicklaunch

        .add(navigationItems[i].title, navigationItems[i].url, true)
        .then(async (res: any) => {
          // console.log("Master Id > ", res.data.Id);

          for (let j: number = 0; navigationItems[i].children.length > j; j++) {
            // await sp.web.navigation.quicklaunch
            await xxweb.navigation.quicklaunch
              .getById(res.data.Id)
              .children.add(
                navigationItems[i].children[j].title,
                navigationItems[i].children[j].url,
                true
              )
              .then(async (child: any) => {
                // console.log("child > ", child);

                for (
                  let k: number = 0;
                  navigationItems[i].children[j].children.length > k;
                  k++
                ) {
                  //   await sp.web.navigation.quicklaunch
                  await xxweb.navigation.quicklaunch
                    .getById(child.data.Id)
                    .children.add(
                      navigationItems[i].children[j].children[k].title,
                      navigationItems[i].children[j].children[k].url,
                      true
                    )
                    .then((subchild: any) => {
                      // console.log("subchild > ", subchild);
                    })
                    .catch((errsubchild: any) => {
                      setIsLoader(false);

                      console.log("errsubchild > ", errsubchild);
                    });

                  if (
                    navigationItems[i].children[j].children.length ===
                    k + 1
                  ) {
                    setIsLoader(false);
                  }
                }
              })
              .catch((errChild: any) => {
                setIsLoader(false);

                console.log("errChild > ", errChild);
              });

            if (navigationItems[i].children.length === j + 1) {
              setIsLoader(false);
            }
          }
        })
        .catch((err: any) => {
          console.log("err > ", err);
          setIsLoader(false);
        });

      if (navigationItems.length === i + 1) {
        setIsLoader(false);
      }
    }
  };
  const getSubsiteName = () => {
    sp.web.webs
      .select("Title", "Url", "Description")
      .get()
      .then((res) => {
        let Titlearray = [];
        res.forEach((val) => {
          Titlearray.push(val.Title);
        });
        // console.log(Titlearray);
        setMasterData([...Titlearray]);
      })
      .catch((err) => {
        setIsLoader(false);
        console.log(err);
      });
  };
  React.useEffect(() => {
    getSubsiteName();
  }, []);
  return (
    <div>
      {isLoader ? (
        <Loader />
      ) : (
        <>
          <Text variant={"xLarge"} style={{ margin: "10px 0" }}>
            Create Release
          </Text>
          <div>
            <p className={styles.lblInfo}>Info:</p>
            <p>
              The administrator can create a new subsite to configure for a new
              release in the product by entering the release name and submitting
              it. This process results in the creation of subsites along with
              pages and navigations. Please ensure that the administrator has
              either Site Collection Administrator or Site Owner permissions for
              the successful creation of the subsite.
            </p>
          </div>
          <div style={{ display: "flex", alignItems: "end", gap: "10px" }}>
            <TextField
              placeholder="Enter the title of new release . . ."
              styles={{
                root: {
                  width: "90%",
                },
              }}
              // errorMessage={error ? error : ""}
              label="Subsite Name"
              onChange={async (e, val) => {
                await getSubsiteName();
                const titleExists = await masterData.some((item) => {
                  return item.toLowerCase().trim() === val.toLowerCase().trim();
                });
                if (val.trim() === "") {
                  await setError("This is required");
                } else if (titleExists) {
                  await setError("This value already exists");
                  // alert("this value already exist !");
                } else {
                  await setError("");
                }

                setValue(val);
              }}
            ></TextField>
            <PrimaryButton
              text="Submit"
              disabled={error ? true : false}
              onClick={(_) => {
                setIsLoader(true);
                !isLoader && SubsiteCreate();
              }}
            />
          </div>
          {error && (
            <p
              style={{
                margin: 0,
                color: "#c00909",
                fontSize: "14px",
                fontWeight: 400,
              }}
            >
              {error}
            </p>
          )}
          {Data.length > 0 && <p className={styles.lblInfo}>Releases:</p>}
          <div className={styles.createdSubSites}>
            {Data.length > 0 &&
              Data.map((val1, i) => {
                return (
                  <div
                    className={styles.createdSiteLabel}
                    onClick={() => {
                      window.open(`${pageurl}/${val1}`, "_blank").focus();
                    }}
                  >
                    {`${i + 1}.${val1} `}
                  </div>
                );
              })}
          </div>
        </>
      )}
    </div>
  );
};
export default MainComponent;
