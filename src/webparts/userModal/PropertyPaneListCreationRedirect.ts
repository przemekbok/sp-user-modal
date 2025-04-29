import { IPropertyPaneField, PropertyPaneFieldType, IPropertyPaneCustomFieldProps } from "@microsoft/sp-property-pane";

export class PropertyPaneListCreationRedirect implements IPropertyPaneField<IPropertyPaneCustomFieldProps> {
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public properties: IPropertyPaneCustomFieldProps;

    constructor(context: any) {
         this.properties = {
             key: "ListCreationRedirect",
             context: context,
             onRender: this.onRender.bind(this)
        };
    }

    private onRender(elem: HTMLElement): void {
        elem.innerHTML = `
        <div style="margin-top: 10px">
            <div>Can't find your list? Create a new one! <a href="${this.properties.context.pageContext.web.absoluteUrl}/_layouts/15/createlist.aspx">Click here</a></div>
        </div>`;
    }
}
export default PropertyPaneListCreationRedirect;