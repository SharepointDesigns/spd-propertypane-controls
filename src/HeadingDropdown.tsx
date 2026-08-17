import * as React from "react";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { Slider } from "@fluentui/react/lib/Slider";

export interface IHeadingDropdownProps {
  selected: string;
  customSize?: number;
  onChange: (val: string) => void;
  onCustomSizeChange?: (val: number) => void;
}

const HeadingDropdown: React.FC<IHeadingDropdownProps> = (props) => {
  const [isCustomSelected, setIsCustomSelected] = React.useState(
    props.selected === "custom"
  );

  const options: IDropdownOption[] = [
    { key: "32", text: "Heading 1" },
    { key: "28", text: "Heading 2" },
    { key: "24", text: "Heading 3" },
    { key: "20", text: "Heading 4" },
    { key: "18", text: "Normal" },
    { key: "custom", text: "Custom Font Size" },
  ];

  const handleDropdownChange = (_: unknown, option?: IDropdownOption): void => {
    if (!option) return;
    props.onChange(option.key as string);
    setIsCustomSelected(option.key === "custom");
  };

  return (
    <div>
      <Dropdown
        label="Choose title heading level"
        selectedKey={props.selected}
        onChange={handleDropdownChange}
        options={options}
        onRenderOption={(option) => {
          switch (option?.key) {
            case "32":
              return (
                <span style={{ fontSize: 32, fontWeight: 600 }}>Heading 1</span>
              );
            case "28":
              return (
                <span style={{ fontSize: 28, fontWeight: 600 }}>Heading 2</span>
              );
            case "24":
              return (
                <span style={{ fontSize: 24, fontWeight: 600 }}>Heading 3</span>
              );
            case "20":
              return (
                <span style={{ fontSize: 20, fontWeight: 600 }}>Heading 4</span>
              );
            case "custom":
              return (
                <span style={{ fontSize: 20, fontWeight: 600 }}>
                  Custom Font Size
                </span>
              );
            default:
              return (
                <span style={{ fontSize: 18, fontWeight: 600 }}>Normal</span>
              );
          }
        }}
      />

      {isCustomSelected && (
        <div style={{ marginTop: 12 }}>
          <Slider
            label="Select Custom Font Size"
            min={10}
            max={68}
            step={2}
            value={props.customSize || 20}
            showValue
            onChange={(value) => props.onCustomSizeChange?.(value)}
          />
        </div>
      )}
    </div>
  );
};

export default HeadingDropdown;
