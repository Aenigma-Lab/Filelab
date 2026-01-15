import { useState, useRef, useEffect, useMemo } from "react";
import axios from "axios";
import { toast } from "sonner";
import { Upload, Download, Trash2, RotateCw, Image as ImageIcon, Type, Settings, Palette, Eye, EyeOff, FileText, AlertCircle } from "lucide-react";
import { Button } from "@/components/ui/button";
import ProgressButton from "@/components/ui/ProgressButton";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Slider } from "@/components/ui/slider";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Checkbox } from "@/components/ui/checkbox";
import { cn } from "@/lib/utils";

// ============== FILE SIZE CONSTANTS ==============
// Maximum file size: 30MB (30 * 1024 * 1024 bytes)
const MAX_FILE_SIZE = 30 * 1024 * 1024;
// Format for displaying file size
const formatFileSize = (bytes) => {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
};

// Create axios instance with interceptors for better error handling
const api = axios.create({
  // Use relative path for network compatibility - nginx proxies /api to backend
  baseURL: process.env.REACT_APP_BACKEND_URL ? `${process.env.REACT_APP_BACKEND_URL}/api` : "/api",
  timeout: 60000, // 60 seconds timeout
  headers: {
    'Content-Type': 'multipart/form-data'
  }
});

// Flag to prevent error interceptor recursion
let isProcessingError = false;

// Add response interceptor for better error handling
api.interceptors.response.use(
  response => response,
  error => {
    // Prevent infinite recursion by marking errors that have already been processed
    if (isProcessingError) {
      // Return the original error without re-processing
      return Promise.reject(error);
    }
    
    isProcessingError = true;
    
    try {
      console.error('[WATERMARK API ERROR]', {
        message: error.message,
        code: error.code,
        status: error.response?.status,
        data: error.response?.data,
        config: {
          url: error.config?.url,
          method: error.config?.method
        }
      });
      
      // Handle specific error cases
      if (error.code === 'ECONNABORTED') {
        return Promise.reject(new Error('Request timed out. The server may be busy. Please try again.'));
      }
      
      if (!error.response) {
        // Network error - server might not be running
        return Promise.reject(new Error('Network error. Please check if the backend server is running.'));
      }
      
      // For HTTP errors, pass through the original error with status info
      const status = error.response?.status;
      let errorMessage = `Server error (${status})`;
      
      if (status === 400) {
        errorMessage = error.response?.data?.detail || 'Invalid request. Please check your input.';
      } else if (status === 413) {
        errorMessage = 'File too large. Please upload a smaller file.';
      } else if (status === 415) {
        errorMessage = 'Unsupported file format. Please upload a valid PDF file.';
      } else if (status === 500) {
        errorMessage = 'Server error. Please try again later.';
      }
      
      return Promise.reject(new Error(errorMessage));
    } finally {
      // Reset the flag after processing
      isProcessingError = false;
    }
  }
);

// Use relative path for network compatibility - nginx proxies /api to backend
const API = process.env.REACT_APP_BACKEND_URL ? `${process.env.REACT_APP_BACKEND_URL}/api` : "/api";

// Health check function to verify backend is running
const checkBackendHealth = async () => {
  try {
    // Use relative path for health check - works from any network location
    const healthUrl = process.env.REACT_APP_BACKEND_URL 
      ? `${process.env.REACT_APP_BACKEND_URL}/api/` 
      : "/api/";
    const response = await fetch(healthUrl);
    return response.ok;
  } catch (error) {
    console.error('Backend health check failed:', error);
    return false;
  }
};

const WatermarkPDF = () => {
  const [loading, setLoading] = useState(false);
  const [success, setSuccess] = useState(false);
  const [conversionProgress, setConversionProgress] = useState(0);
  const [showDownloadButton, setShowDownloadButton] = useState(false);
  const [convertedBlob, setConvertedBlob] = useState(null);
  const [convertedFilename, setConvertedFilename] = useState("");

  // PDF file state
  const [pdfFile, setPdfFile] = useState(null);
  const [watermarkImage, setWatermarkImage] = useState(null);
  const pdfInputRef = useRef(null);
  const watermarkInputRef = useRef(null);

  // Text watermark state
  const [textWatermark, setTextWatermark] = useState({
    text: "CONFIDENTIAL",
    fontName: "Helvetica-Bold",
    fontSize: 48,
    color: "#808080",
    opacity: 0.3,
    rotation: 45,
    position: "center",
    firstPageOnly: false,
    pageRanges: "",
    marginX: 50,
    marginY: 50,
    outline: false,
    outlineColor: "#FFFFFF"
  });

  // Image watermark state
  const [imageWatermark, setImageWatermark] = useState({
    opacity: 0.3,
    position: "center",
    scale: 0.5,
    rotation: 0,
    firstPageOnly: false,
    pageRanges: "",
    marginX: 50,
    marginY: 50
  });

  const [activeTab, setActiveTab] = useState('text'); // Track active watermark tab

  // Preview state
  const [showPreview, setShowPreview] = useState(false);
  const [previewUrl, setPreviewUrl] = useState(null);
  const [imagePreviewUrl, setImagePreviewUrl] = useState(null);

  // Create preview URL when PDF is selected
  useEffect(() => {
    if (pdfFile) {
      const url = URL.createObjectURL(pdfFile);
      setPreviewUrl(url);
      return () => URL.revokeObjectURL(url);
    } else {
      setPreviewUrl(null);
      setShowPreview(false);
    }
  }, [pdfFile]);

  // Create preview URL when watermark image is selected
  useEffect(() => {
    if (watermarkImage) {
      const url = URL.createObjectURL(watermarkImage);
      setImagePreviewUrl(url);
      return () => URL.revokeObjectURL(url);
    } else {
      setImagePreviewUrl(null);
    }
  }, [watermarkImage]);

  // Get text watermark CSS styles
  const getTextWatermarkStyle = useMemo(() => {
    const baseStyle = {
      fontSize: `${textWatermark.fontSize}px`,
      color: textWatermark.color,
      opacity: textWatermark.opacity,
      fontFamily: textWatermark.fontName.replace('-', ' '),
      fontWeight: textWatermark.fontName.includes('Bold') ? 'bold' : 'normal',
      fontStyle: textWatermark.fontName.includes('Oblique') ? 'italic' : 'normal',
      transform: `rotate(${textWatermark.rotation}deg)`,
      '--text-color': textWatermark.color,
      '--outline-color': textWatermark.outlineColor,
    };
    return baseStyle;
  }, [textWatermark]);

  // Get image watermark CSS styles
  const getImageWatermarkStyle = useMemo(() => {
    return {
      opacity: imageWatermark.opacity,
      transform: `rotate(${imageWatermark.rotation}deg) scale(${imageWatermark.scale})`,
      maxWidth: `${imageWatermark.scale * 200}px`,
      maxHeight: `${imageWatermark.scale * 200}px`,
    };
  }, [imageWatermark]);

  // Generate tiled watermarks
  const tiledWatermarks = useMemo(() => {
    return Array.from({ length: 12 }, (_, i) => i);
  }, []);

  // Font options
  const fonts = [
    "Helvetica",
    "Helvetica-Bold",
    "Helvetica-Oblique",
    "Helvetica-BoldOblique",
    "Times-Roman",
    "Times-Bold",
    "Times-Italic",
    "Times-BoldItalic",
    "Courier",
    "Courier-Bold",
    "Courier-Oblique",
    "Courier-BoldOblique"
  ];

  // Position options
  const positions = [
    { value: "center", label: "Center" },
    { value: "top_left", label: "Top Left" },
    { value: "top_right", label: "Top Right" },
    { value: "bottom_left", label: "Bottom Left" },
    { value: "bottom_right", label: "Bottom Right" },
    { value: "tiled", label: "Tiled (Repeated)" }
  ];

  // Color presets
  const colorPresets = [
    "#808080", // Gray
    "#FF0000", // Red
    "#0000FF", // Blue
    "#00FF00", // Green
    "#FFFF00", // Yellow
    "#FF00FF", // Magenta
    "#00FFFF", // Cyan
    "#000000", // Black
    "#FFFFFF", // White
    "#FFA500", // Orange
    "#800080", // Purple
    "#008000"  // Dark Green
  ];

  const handlePdfSelect = (e) => {
    const file = e.target.files[0];
    if (file && file.type === "application/pdf") {
      // Check file size before accepting
      if (file.size > MAX_FILE_SIZE) {
        toast.error(
          <div className="flex items-center gap-2">
            <AlertCircle className="w-4 h-4 text-red-500" />
            <span>File too large! Maximum size is {formatFileSize(MAX_FILE_SIZE)}. Your file is {formatFileSize(file.size)}</span>
          </div>
        );
        if (pdfInputRef.current) {
          pdfInputRef.current.value = '';
        }
        return;
      }
      setPdfFile(file);
      toast.success(`Selected: ${file.name} (${formatFileSize(file.size)})`);
    } else {
      toast.error("Please select a PDF file");
    }
    if (pdfInputRef.current) {
      pdfInputRef.current.value = '';
    }
  };

  const handleWatermarkImageSelect = (e) => {
    const file = e.target.files[0];
    if (file && file.type.startsWith('image/')) {
      setWatermarkImage(file);
      toast.success(`Selected watermark: ${file.name}`);
    } else {
      toast.error("Please select an image file (PNG, JPG, etc.)");
    }
    if (watermarkInputRef.current) {
      watermarkInputRef.current.value = '';
    }
  };

  const clearPdfFile = () => {
    setPdfFile(null);
    setShowDownloadButton(false);
    setConvertedBlob(null);
    setConvertedFilename("");
  };

  const clearWatermarkImage = () => {
    setWatermarkImage(null);
  };

  const downloadFile = (blob, filename) => {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
    toast.success("Watermarked PDF downloaded!");
  };

  const applyTextWatermark = async () => {
    if (!pdfFile) {
      toast.error("Please select a PDF file");
      return;
    }

    if (!textWatermark.text.trim()) {
      toast.error("Please enter watermark text");
      return;
    }

    // Check if backend is running before making the request
    const isBackendRunning = await checkBackendHealth();
    if (!isBackendRunning) {
      toast.error("Backend server is not running. Please start the server and try again.");
      return;
    }

    setLoading(true);
    setSuccess(false);
    setConversionProgress(0);
    setShowDownloadButton(false);
    setConvertedBlob(null);
    setConvertedFilename("");

    // Start progress tracking
    const progressInterval = setInterval(() => {
      setConversionProgress(prev => {
        const increment = prev < 50 ? 10 : prev < 80 ? 5 : prev < 95 ? 2 : 1;
        return Math.min(prev + increment, 95);
      });
    }, 200);

    const formData = new FormData();
    formData.append("file", pdfFile);
    formData.append("text", textWatermark.text);
    formData.append("font_name", textWatermark.fontName);
    formData.append("font_size", textWatermark.fontSize);
    formData.append("color", textWatermark.color);
    formData.append("opacity", textWatermark.opacity);
    formData.append("rotation", textWatermark.rotation);
    formData.append("position", textWatermark.position);
    formData.append("first_page_only", textWatermark.firstPageOnly);
    formData.append("page_ranges", textWatermark.pageRanges || "");
    formData.append("margin_x", textWatermark.marginX);
    formData.append("margin_y", textWatermark.marginY);
    formData.append("outline", textWatermark.outline);
    formData.append("outline_color", textWatermark.outlineColor);

    try {
      const response = await api.post("/watermark/pdf/text", formData, {
        responseType: 'blob'
      });
      
      clearInterval(progressInterval);
      
      // Validate response
      if (!response.data || response.data.size === 0) {
        throw new Error("Empty response received from server");
      }
      
      setConvertedBlob(response.data);
      const originalName = pdfFile.name.replace('.pdf', '');
      setConvertedFilename(`${originalName}_watermarked.pdf`);
      setShowDownloadButton(true);
      setSuccess(true);
      setConversionProgress(100);
      toast.success("Text watermark applied successfully!");
    } catch (error) {
      console.error("Error applying watermark:", error);
      clearInterval(progressInterval);
      
      // Extract error message properly
      let errorMessage = "Failed to apply watermark";
      
      if (error instanceof Error) {
        errorMessage = error.message;
      } else if (error.response?.data?.detail) {
        errorMessage = error.response.data.detail;
      } else if (error.response?.data) {
        // Try to extract from response data
        const dataStr = JSON.stringify(error.response.data);
        if (dataStr && dataStr.length < 500) {
          errorMessage = dataStr;
        }
      } else if (error.message) {
        errorMessage = error.message;
      }
      
      toast.error(errorMessage);
      setSuccess(false);
      setConversionProgress(0);
    } finally {
      setLoading(false);
    }
  };

  const applyImageWatermark = async () => {
    if (!pdfFile) {
      toast.error("Please select a PDF file");
      return;
    }

    if (!watermarkImage) {
      toast.error("Please select a watermark image");
      return;
    }

    // Check if backend is running before making the request
    const isBackendRunning = await checkBackendHealth();
    if (!isBackendRunning) {
      toast.error("Backend server is not running. Please start the server and try again.");
      return;
    }

    setLoading(true);
    setSuccess(false);
    setConversionProgress(0);
    setShowDownloadButton(false);
    setConvertedBlob(null);
    setConvertedFilename("");

    // Start progress tracking
    const progressInterval = setInterval(() => {
      setConversionProgress(prev => {
        const increment = prev < 50 ? 10 : prev < 80 ? 5 : prev < 95 ? 2 : 1;
        return Math.min(prev + increment, 95);
      });
    }, 200);

    const formData = new FormData();
    formData.append("file", pdfFile);
    formData.append("watermark_file", watermarkImage);
    formData.append("opacity", imageWatermark.opacity);
    formData.append("position", imageWatermark.position);
    formData.append("scale", imageWatermark.scale);
    formData.append("rotation", imageWatermark.rotation);
    formData.append("first_page_only", imageWatermark.firstPageOnly);
    formData.append("page_ranges", imageWatermark.pageRanges || "");
    formData.append("margin_x", imageWatermark.marginX);
    formData.append("margin_y", imageWatermark.marginY);

    try {
      const response = await api.post("/watermark/pdf/image", formData, {
        responseType: 'blob'
      });
      
      clearInterval(progressInterval);
      
      // Validate response
      if (!response.data || response.data.size === 0) {
        throw new Error("Empty response received from server");
      }
      
      setConvertedBlob(response.data);
      const originalName = pdfFile.name.replace('.pdf', '');
      setConvertedFilename(`${originalName}_watermarked.pdf`);
      setShowDownloadButton(true);
      setSuccess(true);
      setConversionProgress(100);
      toast.success("Image watermark applied successfully!");
    } catch (error) {
      console.error("Error applying image watermark:", error);
      clearInterval(progressInterval);
      
      // Extract error message properly
      let errorMessage = "Failed to apply image watermark";
      
      if (error instanceof Error) {
        errorMessage = error.message;
      } else if (error.response?.data?.detail) {
        errorMessage = error.response.data.detail;
      } else if (error.response?.data) {
        // Try to extract from response data
        const dataStr = JSON.stringify(error.response.data);
        if (dataStr && dataStr.length < 500) {
          errorMessage = dataStr;
        }
      } else if (error.message) {
        errorMessage = error.message;
      }
      
      toast.error(errorMessage);
      setSuccess(false);
      setConversionProgress(0);
    } finally {
      setLoading(false);
    }
  };

  return (
    <Card className="w-full">
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <Type className="w-5 h-5" />
          Watermark PDF
        </CardTitle>
        <CardDescription>
          Add text or image watermarks to your PDF documents. Protect your documents with custom watermarks.
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        {/* PDF File Selection */}
        <div className="space-y-4">
          <Label className="text-base font-medium">1. Select PDF File</Label>
          <div
            className="border-2 border-dashed border-gray-300 rounded-xl p-6 text-center hover:border-blue-500 transition-colors cursor-pointer bg-gradient-to-br from-blue-50 to-indigo-50"
            onClick={() => pdfInputRef.current?.click()}
          >
            <Upload className="w-8 h-8 mx-auto mb-3 text-gray-400" />
            <p className="text-sm font-medium mb-1">Drop PDF here or click to browse</p>
            <p className="text-xs text-gray-500">Supports: PDF files</p>
            <input
              ref={pdfInputRef}
              type="file"
              accept=".pdf"
              onChange={handlePdfSelect}
              className="hidden"
            />
          </div>

          {pdfFile && (
            <div className="flex items-center justify-between p-3 bg-blue-50 rounded-lg">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-blue-100 rounded">
                  <Type className="w-4 h-4 text-blue-600" />
                </div>
                <div>
                  <p className="text-sm font-medium">{pdfFile.name}</p>
                  <p className="text-xs text-gray-500">
                    {formatFileSize(pdfFile.size)}
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setShowPreview(!showPreview)}
                  className={cn(
                    "transition-all",
                    showPreview ? "bg-blue-100 border-blue-300" : ""
                  )}
                >
                  {showPreview ? (
                    <>
                      <EyeOff className="w-4 h-4 mr-1" />
                      Hide Preview
                    </>
                  ) : (
                    <>
                      <Eye className="w-4 h-4 mr-1" />
                      Show Preview
                    </>
                  )}
                </Button>
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={clearPdfFile}
                  className="text-red-500 hover:text-red-600"
                >
                  <Trash2 className="w-4 h-4" />
                </Button>
              </div>
            </div>
          )}

          {/* CSS-Only Realtime Watermark Preview */}
          {pdfFile && showPreview && (
            <div className="watermark-preview-container">
              <div className="watermark-preview-header">
                <div className="watermark-preview-title">
                  <Eye className="w-4 h-4" />
                  Realtime Preview
                </div>
              </div>

              {/* PDF Viewer with Watermark Overlay */}
              <div className="watermark-pdf-viewer">
                <div className="watermark-pdf-wrapper">
                  {/* PDF Display */}
                  <object
                    data={previewUrl}
                    type="application/pdf"
                    className="watermark-pdf-object"
                  >
                    <div className="watermark-preview-empty">
                      <FileText className="watermark-preview-empty-icon" />
                      <p className="watermark-preview-empty-text">
                        PDF preview not available in this browser
                      </p>
                    </div>
                  </object>

                  {/* Watermark Overlay Container */}
                  <div className="watermark-overlay-container">
                    {/* Text Watermark Overlay */}
                    {activeTab === 'text' && textWatermark.text.trim() && (
                      textWatermark.position === 'tiled' ? (
                        <div className="watermark-tiled-container">
                          {tiledWatermarks.map((index) => (
                            <div key={index} className="watermark-tiled-item">
                              <span
                                className={cn(
                                  "watermark-text",
                                  textWatermark.outline && "watermark-text-with-outline"
                                )}
                                style={getTextWatermarkStyle}
                              >
                                {textWatermark.text}
                              </span>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div
                          className={cn(
                            "watermark-overlay",
                            `watermark-position-${textWatermark.position}`
                          )}
                        >
                          <span
                            className={cn(
                              "watermark-text",
                              textWatermark.outline && "watermark-text-with-outline"
                            )}
                            style={getTextWatermarkStyle}
                          >
                            {textWatermark.text}
                          </span>
                        </div>
                      )
                    )}

                    {/* Image Watermark Overlay */}
                    {activeTab === 'image' && imagePreviewUrl && (
                      imageWatermark.position === 'tiled' ? (
                        <div className="watermark-tiled-container">
                          {tiledWatermarks.map((index) => (
                            <div key={index} className="watermark-tiled-item">
                              <img
                                src={imagePreviewUrl}
                                alt="Watermark"
                                className="watermark-image"
                                style={getImageWatermarkStyle}
                              />
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div
                          className={cn(
                            "watermark-overlay",
                            `watermark-position-${imageWatermark.position}`
                          )}
                        >
                          <img
                            src={imagePreviewUrl}
                            alt="Watermark"
                            className="watermark-image"
                            style={getImageWatermarkStyle}
                          />
                        </div>
                      )
                    )}
                  </div>
                </div>
              </div>

              {/* Preview Info */}
              <div className="watermark-preview-controls">
                <span className="watermark-preview-controls-label">
                  {activeTab === 'text' ? 'Text Watermark' : 'Image Watermark'} - Updates in realtime
                </span>
              </div>
            </div>
          )}
        </div>

        {/* Watermark Type Tabs */}
        <Tabs defaultValue="text" value={activeTab} onValueChange={setActiveTab} className="w-full">
          <TabsList className="grid w-full grid-cols-2">
            <TabsTrigger value="text" className="flex items-center gap-2">
              <Type className="w-4 h-4" />
              Text Watermark
            </TabsTrigger>
            <TabsTrigger value="image" className="flex items-center gap-2">
              <ImageIcon className="w-4 h-4" />
              Image Watermark
            </TabsTrigger>
          </TabsList>

          {/* Text Watermark */}
          <TabsContent value="text" className="space-y-6 mt-6">
            <div className="grid gap-4 md:grid-cols-2">
              {/* Text Input */}
              <div className="col-span-2">
                <Label htmlFor="watermark-text">Watermark Text</Label>
                <Input
                  id="watermark-text"
                  value={textWatermark.text}
                  onChange={(e) => setTextWatermark({ ...textWatermark, text: e.target.value })}
                  placeholder="Enter watermark text"
                  className="mt-1"
                />
              </div>

              {/* Font Selection */}
              <div>
                <Label htmlFor="font-name">Font</Label>
                <Select
                  value={textWatermark.fontName}
                  onValueChange={(value) => setTextWatermark({ ...textWatermark, fontName: value })}
                >
                  <SelectTrigger id="font-name" className="mt-1">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    {fonts.map((font) => (
                      <SelectItem key={font} value={font}>
                        {font}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              {/* Font Size */}
              <div>
                <Label htmlFor="font-size">Font Size: {textWatermark.fontSize}px</Label>
                <Slider
                  id="font-size"
                  min={8}
                  max={200}
                  step={1}
                  value={[textWatermark.fontSize]}
                  onValueChange={([value]) => setTextWatermark({ ...textWatermark, fontSize: value })}
                  className="mt-2"
                />
              </div>

              {/* Color Selection */}
              <div>
                <Label htmlFor="color">Color</Label>
                <div className="flex items-center gap-2 mt-1">
                  <input
                    type="color"
                    id="color"
                    value={textWatermark.color}
                    onChange={(e) => setTextWatermark({ ...textWatermark, color: e.target.value })}
                    className="w-10 h-10 rounded border cursor-pointer"
                  />
                  <Input
                    value={textWatermark.color}
                    onChange={(e) => setTextWatermark({ ...textWatermark, color: e.target.value })}
                    placeholder="#808080"
                    className="flex-1"
                  />
                </div>
                {/* Color Presets */}
                <div className="flex flex-wrap gap-1 mt-2">
                  {colorPresets.map((color) => (
                    <button
                      key={color}
                      type="button"
                      className={`w-6 h-6 rounded border-2 cursor-pointer transition-transform hover:scale-110 ${
                        textWatermark.color === color ? 'border-blue-500 ring-2 ring-blue-200' : 'border-gray-300'
                      }`}
                      style={{ backgroundColor: color }}
                      onClick={() => setTextWatermark({ ...textWatermark, color })}
                      title={color}
                    />
                  ))}
                </div>
              </div>

              {/* Position */}
              <div>
                <Label htmlFor="position">Position</Label>
                <Select
                  value={textWatermark.position}
                  onValueChange={(value) => setTextWatermark({ ...textWatermark, position: value })}
                >
                  <SelectTrigger id="position" className="mt-1">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    {positions.map((pos) => (
                      <SelectItem key={pos.value} value={pos.value}>
                        {pos.label}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              {/* Opacity */}
              <div>
                <Label htmlFor="opacity">Opacity: {Math.round(textWatermark.opacity * 100)}%</Label>
                <Slider
                  id="opacity"
                  min={0.05}
                  max={1}
                  step={0.05}
                  value={[textWatermark.opacity]}
                  onValueChange={([value]) => setTextWatermark({ ...textWatermark, opacity: value })}
                  className="mt-2"
                />
              </div>

              {/* Rotation */}
              <div>
                <Label htmlFor="rotation">Rotation: {textWatermark.rotation}°</Label>
                <Slider
                  id="rotation"
                  min={-180}
                  max={180}
                  step={5}
                  value={[textWatermark.rotation]}
                  onValueChange={([value]) => setTextWatermark({ ...textWatermark, rotation: value })}
                  className="mt-2"
                />
              </div>

              {/* Outline */}
              <div className="col-span-2 flex items-center space-x-2">
                <Checkbox
                  id="outline"
                  checked={textWatermark.outline}
                  onCheckedChange={(checked) => setTextWatermark({ ...textWatermark, outline: checked })}
                />
                <Label htmlFor="outline" className="cursor-pointer">Add outline to text</Label>
              </div>

              {/* First Page Only */}
              <div className="col-span-2 flex items-center space-x-2">
                <Checkbox
                  id="first-page-only"
                  checked={textWatermark.firstPageOnly}
                  onCheckedChange={(checked) => setTextWatermark({ ...textWatermark, firstPageOnly: checked })}
                />
                <Label htmlFor="first-page-only" className="cursor-pointer">Apply to first page only</Label>
              </div>
            </div>

            {/* Apply Button */}
            <ProgressButton
              variant="default"
              loading={loading}
              success={success && showDownloadButton && convertedBlob}
              progress={conversionProgress}
              onClick={applyTextWatermark}
              disabled={loading || !pdfFile}
              loadingText="Applying Watermark..."
              successMessage="Complete!"
              downloadUrl={success && convertedBlob ? window.URL.createObjectURL(convertedBlob) : null}
              downloadFilename={convertedFilename}
              onDownloadComplete={() => {
                setPdfFile(null);
                setShowDownloadButton(false);
                setConvertedBlob(null);
                setConvertedFilename("");
                setSuccess(false);
              }}
            >
              <Type className="w-4 h-4 mr-2" />
              {success && showDownloadButton && convertedBlob ? 'Download Watermarked PDF' : 'Apply Text Watermark'}
            </ProgressButton>
          </TabsContent>

          {/* Image Watermark */}
          <TabsContent value="image" className="space-y-6 mt-6">
            <div className="space-y-4">
              <Label className="text-base font-medium">2. Select Watermark Image (Logo)</Label>
              <div
                className="border-2 border-dashed border-gray-300 rounded-xl p-6 text-center hover:border-purple-500 transition-colors cursor-pointer bg-gradient-to-br from-purple-50 to-pink-50"
                onClick={() => watermarkInputRef.current?.click()}
              >
                <ImageIcon className="w-8 h-8 mx-auto mb-3 text-gray-400" />
                <p className="text-sm font-medium mb-1">Drop watermark image here or click to browse</p>
                <p className="text-xs text-gray-500">Supports: PNG, JPG, WEBP, BMP</p>
                <input
                  ref={watermarkInputRef}
                  type="file"
                  accept="image/*"
                  onChange={handleWatermarkImageSelect}
                  className="hidden"
                />
              </div>

              {watermarkImage && (
                <div className="flex items-center justify-between p-3 bg-purple-50 rounded-lg">
                  <div className="flex items-center gap-3">
                    <div className="p-2 bg-purple-100 rounded">
                      <ImageIcon className="w-4 h-4 text-purple-600" />
                    </div>
                    <div>
                      <p className="text-sm font-medium">{watermarkImage.name}</p>
                      <p className="text-xs text-gray-500">
                        {(watermarkImage.size / 1024).toFixed(2)} KB
                      </p>
                    </div>
                  </div>
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={clearWatermarkImage}
                    className="text-red-500 hover:text-red-600"
                  >
                    <Trash2 className="w-4 h-4" />
                  </Button>
                </div>
              )}
            </div>

            <div className="grid gap-4 md:grid-cols-2">
              {/* Position */}
              <div>
                <Label htmlFor="img-position">Position</Label>
                <Select
                  value={imageWatermark.position}
                  onValueChange={(value) => setImageWatermark({ ...imageWatermark, position: value })}
                >
                  <SelectTrigger id="img-position" className="mt-1">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    {positions.map((pos) => (
                      <SelectItem key={pos.value} value={pos.value}>
                        {pos.label}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              {/* Scale */}
              <div>
                <Label htmlFor="scale">Scale: {Math.round(imageWatermark.scale * 100)}%</Label>
                <Slider
                  id="scale"
                  min={0.1}
                  max={2}
                  step={0.05}
                  value={[imageWatermark.scale]}
                  onValueChange={([value]) => setImageWatermark({ ...imageWatermark, scale: value })}
                  className="mt-2"
                />
              </div>

              {/* Opacity */}
              <div>
                <Label htmlFor="img-opacity">Opacity: {Math.round(imageWatermark.opacity * 100)}%</Label>
                <Slider
                  id="img-opacity"
                  min={0.05}
                  max={1}
                  step={0.05}
                  value={[imageWatermark.opacity]}
                  onValueChange={([value]) => setImageWatermark({ ...imageWatermark, opacity: value })}
                  className="mt-2"
                />
              </div>

              {/* Rotation */}
              <div>
                <Label htmlFor="img-rotation">Rotation: {imageWatermark.rotation}°</Label>
                <Slider
                  id="img-rotation"
                  min={-180}
                  max={180}
                  step={5}
                  value={[imageWatermark.rotation]}
                  onValueChange={([value]) => setImageWatermark({ ...imageWatermark, rotation: value })}
                  className="mt-2"
                />
              </div>

              {/* First Page Only */}
              <div className="col-span-2 flex items-center space-x-2">
                <Checkbox
                  id="img-first-page-only"
                  checked={imageWatermark.firstPageOnly}
                  onCheckedChange={(checked) => setImageWatermark({ ...imageWatermark, firstPageOnly: checked })}
                />
                <Label htmlFor="img-first-page-only" className="cursor-pointer">Apply to first page only</Label>
              </div>
            </div>

            {/* Apply Button */}
            <ProgressButton
              variant="purple"
              loading={loading}
              success={success && showDownloadButton && convertedBlob}
              progress={conversionProgress}
              onClick={applyImageWatermark}
              disabled={loading || !pdfFile || !watermarkImage}
              loadingText="Applying Watermark..."
              successMessage="Complete!"
              downloadUrl={success && convertedBlob ? window.URL.createObjectURL(convertedBlob) : null}
              downloadFilename={convertedFilename}
              onDownloadComplete={() => {
                setPdfFile(null);
                setWatermarkImage(null);
                setShowDownloadButton(false);
                setConvertedBlob(null);
                setConvertedFilename("");
                setSuccess(false);
              }}
            >
              <ImageIcon className="w-4 h-4 mr-2" />
              {success && showDownloadButton && convertedBlob ? 'Download Watermarked PDF' : 'Apply Image Watermark'}
            </ProgressButton>
          </TabsContent>
        </Tabs>
      </CardContent>
    </Card>
  );
};

export default WatermarkPDF;

