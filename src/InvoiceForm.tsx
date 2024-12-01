import React, { useState, useEffect } from 'react';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { TextField, PrimaryButton, Stack, IStackTokens } from '@fluentui/react';

const InvoiceForm: React.FC = () => {
  // Default values
  const defaultBags = 150;
  const defaultRatePerKg = 66;
  const defaultSgst = 2475;
  const defaultCgst = 2475;
  const defaultKgPerBag = 10;
  const defaultVehicleNumber = 'KA-53AB7553';

  // Initialize state with default values
  const [bags, setBags] = useState<number>(defaultBags); // Specify number type
  const [kgPerBag, setKgPerBag] = useState<number>(defaultKgPerBag); // Specify number type
  const [ratePerKg, setRatePerKg] = useState<number>(defaultRatePerKg); // Specify number type
  const [quantity, setQuantity] = useState<number>(defaultBags * defaultKgPerBag); // Specify number type
  const [totalAmount, setTotalAmount] = useState<number>(defaultBags * defaultKgPerBag * defaultRatePerKg); // Specify number type
  const [sgst, setSgst] = useState<number>(defaultSgst); // SGST in rupees
  const [cgst, setCgst] = useState<number>(defaultCgst); // CGST in rupees
  const [finalAmount, setFinalAmount] = useState<number>(defaultBags * defaultKgPerBag * defaultRatePerKg + defaultSgst + defaultCgst);

  // New fields for Vehicle Number, Invoice Date, and Invoice Number
  const [vehicleNumber, setVehicleNumber] = useState(defaultVehicleNumber);
  const [invoiceDate, setInvoiceDate] = useState('');
  const [invoiceNumber, setInvoiceNumber] = useState('');

  // Recalculate values when bags, kgPerBag, ratePerKg, SGST, or CGST change
  useEffect(() => {
    const calculatedQuantity = bags * kgPerBag; // Total quantity in KG
    const calculatedAmount = calculatedQuantity * ratePerKg; // Total amount without taxes
    const calculatedFinalAmount = calculatedAmount + sgst + cgst; // Total after taxes (added SGST and CGST manually)

    // Update state values
    setQuantity(calculatedQuantity);
    setTotalAmount(calculatedAmount);
    setFinalAmount(calculatedFinalAmount);
  }, [bags, kgPerBag, ratePerKg, sgst, cgst]);

  // This function will be called to generate and download the DOCX file
  const handleDownloadWord = async () => {
    try {
      // Fetch the .docx template from the public folder
      const response = await fetch('/invoice-template.docx');
      const arrayBuffer = await response.arrayBuffer();

      // Load the template into PizZip (this will allow us to manipulate the .docx file)
      const zip = new PizZip(arrayBuffer);

      // Create a Docxtemplater instance with the loaded template
      const doc = new Docxtemplater(zip);

      // Set the data to replace the placeholders in the template
      doc.setData({
        vehicleNumber,
        invoiceDate,
        invoiceNumber,
        bags,
        kgPerBag,
        quantity,
        ratePerKg,
        totalAmount,
        sgst,
        cgst,
        finalAmount,
      });

      // Render the document to replace placeholders with actual data
      doc.render();

      // Generate the new document as a blob
      const out = doc.getZip().generate({ type: 'blob' });

      // Create a download link and trigger the download
      const link = document.createElement('a');
      link.href = URL.createObjectURL(out);
      link.download = 'invoice.docx';
      link.click();
    } catch (error) {
      console.error('Error generating Word file:', error);
    }
  };

  // Fluent UI Stack configuration (for spacing)
  const stackTokens: IStackTokens = { childrenGap: 15 };

  return (
    <div style={{ maxWidth: 800, margin: '0 auto', padding: '20px' }}>
      <h1 style={{ textAlign: 'center' }}>Invoice Generator</h1>
      <form>
        <Stack tokens={stackTokens} styles={{ root: { width: '100%' } }}>
          <TextField
            label="Invoice Number"
            value={invoiceNumber}
            onChange={(e, newValue) => setInvoiceNumber(newValue || '')}
            placeholder="Enter Invoice Number"
            required
            styles={{ root: { width: '100%' } }}
          />
          <TextField
            label="Invoice Date"
            type="date"
            value={invoiceDate}
            onChange={(e, newValue) => setInvoiceDate(newValue || '')}
            required
            styles={{ root: { width: '100%' } }}
          />
          <TextField
            label="Vehicle Number"
            value={vehicleNumber}
            onChange={(e, newValue) => setVehicleNumber(newValue || '')}
            placeholder="Enter Vehicle Number"
            required
            styles={{ root: { width: '100%' } }}
          />
          <TextField
            label="Number of Bags"
            type="number"
            value={bags.toString()} // Convert number to string for TextField
            onChange={(e, newValue) => setBags(Number(newValue) || defaultBags)} // Convert string to number
            placeholder="Enter Number of Bags"
            required
            styles={{ root: { width: '100%' } }}
          />
          <TextField
            label="KG per Bag"
            type="number"
            value={kgPerBag.toString()} // Convert number to string for TextField
            onChange={(e, newValue) => setKgPerBag(Number(newValue) || defaultKgPerBag)} // Convert string to number
            placeholder="Enter KG per Bag"
            required
            styles={{ root: { width: '100%' } }}
          />
          <TextField
            label="Rate per KG (Rs.)"
            type="number"
            value={ratePerKg.toString()} // Convert number to string for TextField
            onChange={(e, newValue) => setRatePerKg(Number(newValue) || defaultRatePerKg)} // Convert string to number
            placeholder="Enter Rate per KG"
            required
            styles={{ root: { width: '100%' } }}
          />
          <TextField
            label="Total Quantity (KG)"
            value={quantity.toString()} // Convert number to string for TextField
            disabled
            placeholder="Calculated Total Quantity"
            styles={{ root: { width: '100%' } }}
          />
          <TextField
            label="Total Amount (Rs.)"
            value={totalAmount.toString()} // Convert number to string for TextField
            disabled
            placeholder="Calculated Total Amount"
            styles={{ root: { width: '100%' } }}
          />
          <TextField
            label="SGST (Rs.)"
            type="number"
            value={sgst.toString()} // Convert number to string for TextField
            onChange={(e, newValue) => setSgst(Number(newValue) || defaultSgst)} // Convert string to number
            placeholder="Enter SGST in Rupees"
            required
            styles={{ root: { width: '100%' } }}
          />
          <TextField
            label="CGST (Rs.)"
            type="number"
            value={cgst.toString()} // Convert number to string for TextField
            onChange={(e, newValue) => setCgst(Number(newValue) || defaultCgst)} // Convert string to number
            placeholder="Enter CGST in Rupees"
            required
            styles={{ root: { width: '100%' } }}
          />
          <TextField
            label="Final Amount (Rs.)"
            value={finalAmount.toString()} // Convert number to string for TextField
            disabled
            placeholder="Total Amount after Taxes"
            styles={{ root: { width: '100%' } }}
          />
          <PrimaryButton text="Generate Word Document" onClick={handleDownloadWord} />
        </Stack>
      </form>
    </div>
  );
};

export default InvoiceForm;
