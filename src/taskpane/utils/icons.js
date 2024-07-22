import {
  bundleIcon,
  Save20Regular,
  Save20Filled,
  DocumentEdit20Filled,
  DocumentEdit20Regular,
  Record20Regular,
  Record20Filled,
  RecordStop20Regular,
  RecordStop20Filled,
  AddSquare20Filled,
  AddSquare20Regular,
  Delete20Filled,
  Delete20Regular,
} from "@fluentui/react-icons";

const SaveIcon = bundleIcon(Save20Filled, Save20Regular);
const EditIcon = bundleIcon(DocumentEdit20Filled, DocumentEdit20Regular);
const RecordStart = bundleIcon(Record20Filled, Record20Regular);
const RecordStop = bundleIcon(RecordStop20Filled, RecordStop20Regular);
const PlusIcon = bundleIcon(AddSquare20Filled, AddSquare20Regular);
const DeleteIcon = bundleIcon(Delete20Filled, Delete20Regular);

export { SaveIcon, EditIcon, RecordStart, RecordStop, PlusIcon, DeleteIcon };
