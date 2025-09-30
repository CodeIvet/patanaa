import { Card, CardHeader } from "@fluentui/react-components";
import { Link, Text } from "@fluentui/react-components";
import logo from "./custom/lib/patanalogo1.png"; // Import logo

export default function About() {
  return (
    <div style={{ display: "flex", justifyContent: "center", padding: "20px" }}>
      <Card style={{ maxWidth: "500px", textAlign: "center", padding: "20px" }}>
        <CardHeader
          image={
            <img
              src={logo}
              alt="Patanaa Logo"
              style={{ width: 64, height: 64, borderRadius: "8px" }}
            />
          }
          header={
            <Text weight="bold" size={600}>
              Patanaa
            </Text>
          }
          description="The Teams App for Board Meetings"
        />
        <div style={{ marginTop: "16px" }}>
          <Text>
            Patana provides a secure and efficient way to manage your board meetings
            within Microsoft Teams.
          </Text>
          <br />
          <Text>
            Developed by the brave people of <br />
            <Link
              href="https://moveoffice.sharepoint.com/sites/b71/SitePages/Customer-%26-Core-Services.aspx#collaboration-automation-solutions"
              target="_blank"
            >
              Collaboration & Automation Solutions
            </Link>
          </Text>
        </div>
      </Card>
    </div>
  );
}
