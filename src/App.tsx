import { useState, useRef, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import pptxgen from 'pptxgenjs';
import {
  Sprout,
  Users,
  Cpu,
  Package,
  Lightbulb,
  GraduationCap,
  Database,
  Briefcase,
  Phone,
  Mail,
  MapPin,
  Calendar,
  Printer,
  Copy,
  Volume2,
  VolumeX,
  CheckCircle,
  BookOpen,
  Award,
  Check,
  ExternalLink,
  ChevronRight,
  User,
  Clock,
  Sparkles,
  FileText,
  AlertCircle,
  X,
  Download,
  DownloadCloud
} from 'lucide-react';
import { THEMES, SUGGESTIONS, BOTANICAL_CROPS, BIHAR_AGRI_HIGHLIGHTS } from './data';
import { ThemeConfig, Suggestion, BotanicalInfo } from './types';

export default function App() {
  const [selectedThemeId, setSelectedThemeId] = useState<'green' | 'golden'>('green');
  const [activeSuggestion, setActiveSuggestion] = useState<number | null>(null);
  const [activeCropIndex, setActiveCropIndex] = useState<number>(0);
  const [isSpeaking, setIsSpeaking] = useState(false);
  const [copied, setCopied] = useState(false);
  
  // Interactive Presentation Deck States
  const [showPresenterMode, setShowPresenterMode] = useState(false);
  const [currentSlideIndex, setCurrentSlideIndex] = useState(0);
  
  // Interactive Minister's Desk State
  const [ministerDecision, setMinisterDecision] = useState<'pending' | 'approved' | 'scheduled'>('pending');
  const [meetingDate, setMeetingDate] = useState('2026-07-15');
  const [meetingTime, setMeetingTime] = useState('11:30');
  const [ministerNote, setMinisterNote] = useState('प्रिय रवि, आपका सुझाव पत्र प्राप्त हुआ। आपके वैज्ञानिक और आधुनिक दृष्टिकोण से मैं अत्यंत प्रभावित हूँ। शीघ्र ही आपसे मुलाकात होगी।');
  const [showApprovalNotice, setShowApprovalNotice] = useState(false);

  const utteranceRef = useRef<SpeechSynthesisUtterance | null>(null);

  const theme: ThemeConfig = THEMES.find(t => t.id === selectedThemeId) || THEMES[0];

  // Raw letter copy text
  const rawLetterText = `सेवा में,
माननीय कृषि मंत्री महोदय,
बिहार सरकार, पटना।

विषय: बिहार में कृषि, किसानों की आय एवं युवा कृषि-उद्यमिता को बढ़ावा देने हेतु सुझाव एवं मुलाकात का अनुरोध।

महोदय,

सादर प्रणाम।

मेरा नाम रवि कुमार है। मैं हिंदुनी, आलमपुर गोनपुरा, फुलवारी शरीफ, पटना, बिहार का निवासी हूँ। मेरी आयु 17 वर्ष 11 माह है। मैंने विज्ञान विषय से 12वीं कक्षा उत्तीर्ण की है और वर्तमान में पटलिपुत्र विश्वविद्यालय से संबद्ध बी.एस. कॉलेज, दानापुर में बी.एससी. (वनस्पति विज्ञान) की पढ़ाई कर रहा हूँ। इस कारण मुझे पौधों, कृषि और प्राकृतिक संसाधनों के बारे में प्रारंभिक वैज्ञानिक जानकारी प्राप्त हो रही है।

मैं एक किसान परिवार से हूँ। मैंने अपने घर और गाँव में देखा है कि कई परिवारों ने खेती छोड़ दी है या केवल जीवन-यापन के लिए खेती कर रहे हैं। किसान कड़ी मेहनत करते हैं, लेकिन उन्हें उनकी मेहनत के अनुसार पर्याप्त लाभ नहीं मिल पाता। बढ़ती लागत, उचित बाजार की कमी, भंडारण सुविधाओं का अभाव और फसलों के उचित मूल्य न मिलने के कारण खेती युवाओं के लिए आकर्षक नहीं रह गई है।

इसी वजह से आज अधिकांश युवा खेती से दूर होकर केवल सरकारी नौकरियों को ही सुरक्षित भविष्य मान रहे हैं। मेरा मानना है कि यदि कृषि को आधुनिक तकनीक, शिक्षा, प्रोसेसिंग उद्योग, डिजिटल मार्केटिंग और उद्यमिता से जोड़ा जाए, तो कृषि फिर से युवाओं के लिए सम्मानजनक और लाभदायक क्षेत्र बन सकती है तथा बिहार की अर्थव्यवस्था को नई दिशा मिल सकती है।

मेरे कुछ सुझाव निम्नलिखित हैं:

1. प्रत्येक जिले में Young Agri Entrepreneur Clubs की स्थापना की जाए, जहाँ युवाओं को आधुनिक खेती और कृषि-उद्यमिता का प्रशिक्षण दिया जाए।

2. युवाओं और किसानों को AI, डिजिटल मार्केटिंग, ड्रोन तकनीक और आधुनिक कृषि पद्धतियों की जानकारी एवं प्रशिक्षण उपलब्ध कराया जाए।

3. मखाना, मक्का, लीची, सब्जियों और अन्य स्थानीय उत्पादों की प्रोसेसिंग, पैकेजिंग और ब्रांडिंग को बढ़ावा दिया जाए, ताकि किसानों को अधिक लाभ मिल सके।

4. कृषि आधारित स्टार्टअप्स के लिए विशेष अनुदान, मार्गदर्शन और इनक्यूबेशन सुविधाएँ उपलब्ध कराई जाएँ।

5. स्कूलों और कॉलेजों में कृषि-उद्यमिता, एग्री-टेक और आधुनिक खेती से संबंधित कार्यक्रम आयोजित किए जाएँ, जिससे युवा कृषि को एक करियर और व्यवसाय के रूप में अपनाने के लिए प्रेरित हों।

6. छोटे किसानों के लिए कोल्ड स्टोरेज, वेयरहाउस, बेहतर बाजार व्यवस्था और उचित मूल्य सुनिश्चित करने की दिशा में विशेष योजनाएँ लागू की जाएँ।

7. युवाओं को केवल नौकरी खोजने वाला नहीं, बल्कि कृषि के माध्यम से रोजगार देने वाला उद्यमी बनने के लिए प्रोत्साहित किया जाए।

मेरा विश्वास है कि यदि सरकार, किसान और युवा मिलकर कार्य करें, तो कृषि बिहार की GDP, रोजगार और समग्र विकास का सबसे मजबूत आधार बन सकती है। मेरा उद्देश्य बिहार के किसानों और युवाओं के लिए सकारात्मक बदलाव लाने में अपना योगदान देना है।

अतः आपसे विनम्र अनुरोध है कि मुझे अपने विचारों और सुझावों पर चर्चा करने का अवसर प्रदान करने की कृपा करें। आपके मार्गदर्शन और सहयोग से बिहार में कृषि और कृषि-उद्यमिता को नई दिशा मिल सकती है।

सधन्यवाद,

रवि कुमार
बी.एससी. (वनस्पति विज्ञान) छात्र, बी.एस. कॉलेज, दानापुर, पटलिपुत्र विश्वविद्यालय
पता: हिंदुनी, आलमपुर गोनपुरा, फुलवारी शरीफ, पटना, बिहार
मोबाइल: 6203023603
ईमेल: kavikumar11825@gmail.com
दिनांक: 02 जुलाई 2026`;

  // Speech Synthesizer for Hindi TTS
  const handleSpeak = () => {
    if (isSpeaking) {
      window.speechSynthesis.cancel();
      setIsSpeaking(false);
      return;
    }

    const cleanText = `
      सेवा में, माननीय कृषि मंत्री महोदय, बिहार सरकार, पटना। 
      विषय, बिहार में कृषि, किसानों की आय एवं युवा कृषि उद्यमिता को बढ़ावा देने हेतु सुझाव।
      सादर प्रणाम। मेरा नाम रवि कुमार है। मैं दानापुर के बी एस कॉलेज में बी एस सी वनस्पति विज्ञान का छात्र हूँ।
      मैंने बिहार में कृषि और युवा उद्यमिता को बढ़ावा देने के लिए सात प्रमुख सुझाव प्रस्तुत किए हैं।
      पहला सुझाव, हर जिले में यंग एग्री एंटरप्रेन्योर क्लब्स की स्थापना।
      दूसरा सुझाव, किसानों को ड्रोन तकनीक, ए आई और डिजिटल मार्केटिंग का प्रशिक्षण।
      तीसरा सुझाव, मखाना, लीची और मक्का की प्रोसेसिंग और ब्रांडिंग।
      चौथा सुझाव, कृषि स्टार्टअप्स के लिए अनुदान और इनक्यूबेशन।
      पांचवा सुझाव, स्कूलों और कॉलेजों में एग्री-टेक कार्यक्रम।
      छठा सुझाव, छोटे किसानों के लिए कोल्ड स्टोरेज और बेहतर बाजार।
      सातवां सुझाव, युवाओं को रोजगार प्रदाता उद्यमी बनने के लिए प्रोत्साहन।
      आशा है आप इस पत्र पर ध्यान देंगे। धन्यवाद, रवि कुमार।
    `;

    const utterance = new SpeechSynthesisUtterance(cleanText);
    utterance.lang = 'hi-IN';
    utterance.rate = 0.9; // clear, professional pace
    utterance.pitch = 1.0;

    utterance.onend = () => setIsSpeaking(false);
    utterance.onerror = () => setIsSpeaking(false);

    utteranceRef.current = utterance;
    setIsSpeaking(true);
    window.speechSynthesis.speak(utterance);
  };

  useEffect(() => {
    return () => {
      window.speechSynthesis.cancel();
    };
  }, []);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (!showPresenterMode) return;
      if (e.key === 'ArrowRight') {
        setCurrentSlideIndex(prev => Math.min(prev + 1, 11));
      } else if (e.key === 'ArrowLeft') {
        setCurrentSlideIndex(prev => Math.max(prev - 1, 0));
      } else if (e.key === 'Escape') {
        setShowPresenterMode(false);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [showPresenterMode]);

  const handleCopyText = async () => {
    try {
      await navigator.clipboard.writeText(rawLetterText);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error("Failed to copy", err);
    }
  };

  const handlePrint = () => {
    window.print();
  };

  const handleDownloadTxt = () => {
    const element = document.createElement("a");
    const file = new Blob([rawLetterText], {type: 'text/plain;charset=utf-8'});
    element.href = URL.createObjectURL(file);
    element.download = "Ravi_Kumar_Bihar_Agri_Suggestions.txt";
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
  };

  const handleDownloadHtml = () => {
    const htmlTemplate = `<!DOCTYPE html>
<html lang="hi">
<head>
  <meta charset="UTF-8">
  <title>बिहार कृषि सुझाव पत्र - रवि कुमार</title>
  <style>
    body { font-family: 'Georgia', serif; padding: 40px; background-color: #fcfbf7; color: #1e293b; line-height: 1.8; max-width: 800px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 8px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
    h1 { text-align: center; color: #065f46; font-size: 24px; border-bottom: 2px solid #065f46; padding-bottom: 15px; margin-bottom: 30px; }
    .subject { background-color: #f0fdf4; border: 1px solid #bbf7d0; padding: 15px; border-radius: 6px; font-weight: bold; margin: 20px 0; }
    .footer { margin-top: 50px; border-top: 1px dashed #cbd5e1; padding-top: 20px; display: flex; justify-content: space-between; }
    .suggestions { margin-left: 20px; }
    .suggestions p { margin-bottom: 12px; }
  </style>
</head>
<body>
  <h1>बिहार कृषि संवर्धन एवं युवा उद्यमिता प्रस्ताव</h1>
  <p><strong>सेवा में,</strong><br>माननीय कृषि मंत्री महोदय,<br>बिहार सरकार, पटना।</p>
  <div class="subject">विषय: बिहार में कृषि, किसानों की आय एवं युवा कृषि-उद्यमिता को बढ़ावा देने हेतु सुझाव एवं मुलाकात का अनुरोध।</div>
  <p><strong>महोदय,</strong><br>सादर प्रणाम।</p>
  <p>मेरा नाम <strong>रवि कुमार</strong> है। मैं हिंदुनी, आलमपुर गोनपुरा, फुलवारी शरीफ, पटना, बिहार का निवासी हूँ। मेरी आयु 17 वर्ष 11 माह है। मैंने विज्ञान विषय से 12वीं कक्षा उत्तीर्ण की है और वर्तमान में पाटलिपुत्र विश्वविद्यालय से संबद्ध बी.एस. कॉलेज, दानापुर में बी.एससी. (वनस्पति विज्ञान) की पढ़ाई कर रहा हूँ। इस कारण मुझे पौधों, कृषि और प्राकृतिक संसाधनों के बारे में प्रारंभिक वैज्ञानिक जानकारी प्राप्त हो रही है।</p>
  <p>मैं एक किसान परिवार से हूँ। मैंने अपने घर और गाँव में देखा है कि कई परिवारों ने खेती छोड़ दी है या केवल जीवन-यापन के लिए खेती कर रहे हैं। किसान कड़ी मेहनत करते हैं, लेकिन उन्हें उनकी मेहनत के अनुसार पर्याप्त लाभ नहीं मिल पाता। बढ़ती लागत, उचित बाजार की कमी, भंडारण सुविधाओं का अभाव और फसलों के उचित मूल्य न मिलने के कारण खेती युवाओं के लिए आकर्षक नहीं रह गई है।</p>
  <p>इसी वजह से आज अधिकांश युवा खेती से दूर होकर केवल सरकारी नौकरियों को ही सुरक्षित भविष्य मान रहे हैं। मेरा मानना है कि यदि कृषि को आधुनिक तकनीक, शिक्षा, प्रोसेसिंग उद्योग, डिजिटल मार्केटिंग और उद्यमिता से जोड़ा जाए, तो कृषि फिर से युवाओं के लिए सम्मानजनक और लाभदायक क्षेत्र बन सकती है तथा बिहार की अर्थव्यवस्था को नई दिशा मिल सकती है।</p>
  <p><strong>मेरे कुछ सुझाव निम्नलिखित हैं:</strong></p>
  <div class="suggestions">
    <p><strong>1.</strong> प्रत्येक जिले में Young Agri Entrepreneur Clubs की स्थापना की जाए, जहाँ युवाओं को आधुनिक खेती और कृषि-उद्यमिता का प्रशिक्षण दिया जाए।</p>
    <p><strong>2.</strong> युवाओं और किसानों को AI, डिजिटल मार्केटिंग, ड्रोन तकनीक और आधुनिक कृषि पद्धतियों की जानकारी एवं प्रशिक्षण उपलब्ध कराया जाए।</p>
    <p><strong>3.</strong> मखाना, मक्का, लीची, सब्जियों और अन्य स्थानीय उत्पादों की प्रोसेसिंग, पैकेजिंग और ब्रांडिंग को बढ़ावा दिया जाए, ताकि किसानों को अधिक लाभ मिल सके।</p>
    <p><strong>4.</strong> कृषि आधारित स्टार्टअप्स के लिए विशेष अनुदान, मार्गदर्शन और इनक्यूबेशन सुविधाएँ उपलब्ध कराई जाएँ।</p>
    <p><strong>5.</strong> स्कूलों और कॉलेजों में कृषि-उद्यमिता, एग्री-टेक और आधुनिक खेती से संबंधित कार्यक्रम आयोजित किए जाएँ, जिससे युवा कृषि को एक करियर और व्यवसाय के रूप में अपनाने के लिए प्रेरित हों।</p>
    <p><strong>6.</strong> छोटे किसानों के लिए कोल्ड स्टोरेज, वेयरहाउस, बेहतर बाजार व्यवस्था और उचित मूल्य सुनिश्चित करने की दिशा में विशेष योजनाएँ लागू की जाएँ।</p>
    <p><strong>7.</strong> युवाओं को केवल नौकरी खोजने वाला नहीं, बल्कि कृषि के माध्यम से रोजगार देने वाला उद्यमी बनने के लिए प्रोत्साहित किया जाए।</p>
  </div>
  <p>मेरा विश्वास है कि यदि सरकार, किसान और युवा मिलकर कार्य करें, तो कृषि बिहार की GDP, रोजगार और समग्र विकास का सबसे मजबूत आधार बन सकती है। मेरा उद्देश्य बिहार के किसानों और युवाओं के लिए सकारात्मक बदलाव लाने में अपना योगदान देना है।</p>
  <p>अतः आपसे विनम्र अनुरोध है कि मुझे अपने विचारों और सुझावों पर चर्चा करने का अवसर प्रदान करने की कृपा करें। आपके मार्गदर्शन और सहयोग से बिहार में कृषि और कृषि-उद्यमिता को नई दिशा मिल सकती है।</p>
  <div class="footer">
    <div>
      <p>दिनांक: 02 जुलाई 2026</p>
      <p>स्थान: पटना, बिहार</p>
    </div>
    <div style="text-align: right;">
      <p>सधन्यवाद,</p>
      <p><strong>रवि कुमार</strong></p>
      <p>बी.एससी. (वनस्पति विज्ञान) छात्र, बी.एस. कॉलेज, दानापुर</p>
      <p>मोबाइल: 6203023603 | ईमेल: kavikumar11825@gmail.com</p>
    </div>
  </div>
</body>
</html>`;
    const element = document.createElement("a");
    const file = new Blob([htmlTemplate], {type: 'text/html;charset=utf-8'});
    element.href = URL.createObjectURL(file);
    element.download = "Ravi_Kumar_Bihar_Agri_Suggestions.html";
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
  };

  const handleDownloadPpt = () => {
    const pptx = new pptxgen();
    
    // Set presentation to 16:9 widescreen
    pptx.layout = 'LAYOUT_16X9';
    
    // Colors based on selected theme
    const primaryColor = selectedThemeId === 'green' ? '065F46' : 'B45309';
    const secondaryColor = selectedThemeId === 'green' ? 'D1FAE5' : 'FEF3C7';
    const textColor = '1E293B';
    const lightBg = 'F9FAFB';

    // Slide 1: Title Slide
    const slide1 = pptx.addSlide();
    slide1.background = { fill: secondaryColor };
    
    slide1.addText("बिहार कृषि संवर्धन एवं युवा उद्यमिता प्रस्ताव", {
      x: 0.5, y: 1.5, w: 12.33, h: 1.2,
      fontSize: 28, bold: true, color: primaryColor,
      align: 'center', fontFace: 'Arial'
    });
    
    slide1.addText("माननीय कृषि मंत्री महोदय, बिहार सरकार के समक्ष प्रस्तुत नीतिगत सुझाव", {
      x: 0.5, y: 2.8, w: 12.33, h: 0.6,
      fontSize: 16, color: '475569',
      align: 'center', fontFace: 'Arial'
    });

    slide1.addText("प्रस्तावक: रवि कुमार\nबी.एससी. (वनस्पति विज्ञान) छात्र, बी.एस. कॉलेज, दानापुर, पाटलिपुत्र विश्वविद्यालय\nपता: हिंदुनी, आलमपुर गोनपुरा, फुलवारी शरीफ, पटना, बिहार | मो: 6203023603\nदिनांक: 02 जुलाई 2026", {
      x: 0.5, y: 4.0, w: 12.33, h: 2.0,
      fontSize: 13, color: textColor,
      align: 'center', fontFace: 'Arial', lineSpacing: 22
    });

    // Slide 2: Proposer Profile & Motivation
    const slide2 = pptx.addSlide();
    slide2.background = { fill: 'FFFFFF' };
    
    slide2.addText("प्रस्तावक परिचय एवं प्रेरणा (Proposer Profile)", {
      x: 0.5, y: 0.4, w: 12.33, h: 0.6,
      fontSize: 22, bold: true, color: primaryColor, fontFace: 'Arial'
    });

    slide2.addText("रवि कुमार — वनस्पति विज्ञान छात्र", {
      x: 0.5, y: 1.2, w: 5.5, h: 0.4,
      fontSize: 16, bold: true, color: '0F172A', fontFace: 'Arial'
    });

    slide2.addText([
      { text: "• छात्र विवरण: ", options: { bold: true, color: primaryColor } },
      { text: "पाटलिपुत्र विश्वविद्यालय से संबद्ध बी.एस. कॉलेज, दानापुर में बी.एससी. (वनस्पति विज्ञान) के छात्र। आयु: 17 वर्ष 11 माह।\n\n" },
      { text: "• पारिवारिक पृष्ठभूमि: ", options: { bold: true, color: primaryColor } },
      { text: "एक पारंपरिक प्रगतिशील किसान परिवार से संबंध। बचपन से ही कृषि के जमीनी स्तर के संघर्षों और संभावनाओं का साक्षात अनुभव।\n\n" },
      { text: "• वैज्ञानिक प्रेरणा: ", options: { bold: true, color: primaryColor } },
      { text: "वनस्पति विज्ञान (Botany) का छात्र होने के कारण पौधों के जैव-विकास, मृदा विज्ञान और कृषि उत्पादों के मूल्यवर्धन (value-addition) की प्रारंभिक वैज्ञानिक समझ उपलब्ध है।" }
    ], {
      x: 0.5, y: 1.8, w: 5.5, h: 4.5,
      fontSize: 12.5, color: textColor, fontFace: 'Arial', lineSpacing: 20
    });

    slide2.addText("मुख्य ध्येय एवं संकल्प (The Mission)", {
      x: 6.5, y: 1.2, w: 5.5, h: 0.4,
      fontSize: 16, bold: true, color: '0F172A', fontFace: 'Arial'
    });

    slide2.addText([
      { text: "• कृषि से युवाओं का पलायन: ", options: { bold: true, color: '991B1B' } },
      { text: "आज के युवा पारंपरिक खेती को अलाभकारी मानकर पलायन कर रहे हैं। इस पलायन को रोकना और खेती को एक गरिमामयी व्यवसाय बनाना अत्यंत आवश्यक है।\n\n" },
      { text: "• कृषि का आधुनिकीकरण: ", options: { bold: true, color: primaryColor } },
      { text: "मखाना, लीची, मक्का और आम जैसी फसलों में बिहार देश में अग्रणी है। डिजिटल तकनीक (ड्रोन, AI) और सुदृढ़ प्रोसेसिंग उद्योगों के बिना किसानों की आय को दोगुना करना असंभव है।\n\n" },
      { text: "• आत्मनिर्भर बिहार: ", options: { bold: true, color: primaryColor } },
      { text: "बिहार के शिक्षित युवाओं को रोजगार की तलाश करने वालों (Job Seekers) से रोजगार प्रदाता (Job Creators) के रूप में रूपांतरित करना ही हमारा परम लक्ष्य है।" }
    ], {
      x: 6.5, y: 1.8, w: 5.8, h: 4.5,
      fontSize: 12.5, color: textColor, fontFace: 'Arial', lineSpacing: 20
    });

    // Slides 3 to 9: Suggestion Details
    SUGGESTIONS.forEach((s) => {
      const slide = pptx.addSlide();
      slide.background = { fill: lightBg };

      // Slide Title
      slide.addText(`सुझाव #${s.id}: ${s.title}`, {
        x: 0.5, y: 0.4, w: 12.33, h: 0.6,
        fontSize: 20, bold: true, color: primaryColor, fontFace: 'Arial'
      });

      // Original text from letter
      slide.addText("प्रस्ताव का मूल स्वरूप (The Recommendation):", {
        x: 0.5, y: 1.2, w: 11.5, h: 0.3,
        fontSize: 13, bold: true, color: '475569', fontFace: 'Arial'
      });

      slide.addText(`"${s.originalText}"`, {
        x: 0.5, y: 1.6, w: 11.5, h: 1.0,
        fontSize: 14, italic: true, color: primaryColor,
        bg: secondaryColor, fontFace: 'Arial', padding: 12
      });

      // Detailed Context / Scientific analysis
      slide.addText("विस्तृत योजना एवं वैज्ञानिक लाभ (Execution & Impact Analysis):", {
        x: 0.5, y: 2.8, w: 11.5, h: 0.3,
        fontSize: 13, bold: true, color: '475569', fontFace: 'Arial'
      });

      slide.addText(s.detailedContext, {
        x: 0.5, y: 3.2, w: 11.5, h: 2.2,
        fontSize: 12.5, color: textColor, fontFace: 'Arial', lineSpacing: 22
      });

      // Target/Stat Box
      if (s.stats) {
        slide.addText(`विशिष्ट लक्ष्य / लक्षित लाभ: ${s.stats}`, {
          x: 0.5, y: 5.6, w: 11.5, h: 0.6,
          fontSize: 13, bold: true, color: 'FFFFFF',
          bg: primaryColor, fontFace: 'Arial', align: 'center'
        });
      }
    });

    // Slide 10: Bihar Crop Lab / Botanical value addition
    const slide10 = pptx.addSlide();
    slide10.background = { fill: 'FFFFFF' };

    slide10.addText("बिहार के प्रमुख उत्पादों में वैज्ञानिक मूल्यवर्धन (Botanical Value-Addition)", {
      x: 0.5, y: 0.4, w: 12.33, h: 0.6,
      fontSize: 20, bold: true, color: primaryColor, fontFace: 'Arial'
    });

    slide10.addText("वनस्पति विज्ञान छात्र के दृष्टिकोण से बिहार के विशिष्ट उत्पादों की प्रसंस्करण क्षमता:", {
      x: 0.5, y: 1.1, w: 11.5, h: 0.4,
      fontSize: 13, bold: true, color: '475569', fontFace: 'Arial'
    });

    BOTANICAL_CROPS.forEach((crop, idx) => {
      const colX = 0.5 + (idx * 3.0);
      slide10.addText([
        { text: `${crop.hindiName}\n`, options: { bold: true, fontSize: 13, color: primaryColor } },
        { text: `${crop.scientificName}\n`, options: { italic: true, fontSize: 10, color: '475569' } },
        { text: `${crop.family}\n\n`, options: { fontSize: 10, color: '64748B' } },
        { text: `${crop.description.slice(0, 150)}...`, options: { fontSize: 11, color: textColor } }
      ], {
        x: colX, y: 1.7, w: 2.8, h: 4.5,
        bg: 'F8FAFC', border: { type: 'line', color: 'E2E8F0', size: 1 },
        padding: 10, fontFace: 'Arial', lineSpacing: 18
      });
    });

    // Slide 11: Conclusion & Meeting Request
    const slide11 = pptx.addSlide();
    slide11.background = { fill: secondaryColor };

    slide11.addText("सादर अनुरोध एवं मार्गदर्शन की प्रार्थना (Meeting Request)", {
      x: 0.5, y: 1.2, w: 12.33, h: 0.8,
      fontSize: 24, bold: true, color: primaryColor,
      align: 'center', fontFace: 'Arial'
    });

    slide11.addText("माननीय कृषि मंत्री महोदय से विनम्र अनुरोध है कि उपर्युक्त सात क्रांतिकारी प्रस्तावों पर चर्चा करने, इनका क्रियान्वयन सुनिश्चित करने और बिहार के युवा कृषि-छात्रों का उत्साहवर्धन करने के लिए मुझे 10 मिनट की मुलाकात एवं अमूल्य मार्गदर्शन प्रदान करने की कृपा करें।", {
      x: 1.0, y: 2.2, w: 11.33, h: 1.8,
      fontSize: 15, color: textColor,
      align: 'center', fontFace: 'Arial', lineSpacing: 26
    });

    slide11.addText("रवि कुमार\nबी.एससी. (वनस्पति विज्ञान) छात्र, बी.एस. कॉलेज, दानापुर, पाटलिपुत्र विश्वविद्यालय\nमो: +91 6203023603 | ईमेल: kavikumar11825@gmail.com\nपता: हिंदुनी, आलमपुर गोनपुरा, फुलवारी शरीफ, पटना, बिहार", {
      x: 0.5, y: 4.2, w: 12.33, h: 1.8,
      fontSize: 13, bold: true, color: '475569',
      align: 'center', fontFace: 'Arial', lineSpacing: 20
    });

    // Trigger Download
    pptx.writeFile({ fileName: "Ravi_Kumar_Bihar_Agriculture_Suggestions_Pitch_Deck.pptx" });
  };

  // Helper to resolve icon components dynamically
  const renderIcon = (iconName: string, className = "w-5 h-5") => {
    const iconsMap: Record<string, any> = {
      Users: <Users className={className} />,
      Cpu: <Cpu className={className} />,
      Package: <Package className={className} />,
      Lightbulb: <Lightbulb className={className} />,
      GraduationCap: <GraduationCap className={className} />,
      Database: <Database className={className} />,
      Briefcase: <Briefcase className={className} />,
      Sprout: <Sprout className={className} />,
      Droplet: <Sprout className={`${className} text-blue-500`} />,
      Cherry: <Award className={`${className} text-red-500`} />,
      Corn: <Sprout className={`${className} text-yellow-500`} />,
      Sun: <Award className={`${className} text-amber-500`} />
    };
    return iconsMap[iconName] || <FileText className={className} />;
  };

  return (
    <div 
      className="min-h-screen bg-cover bg-center bg-fixed font-sans selection:bg-emerald-200 transition-all duration-700 ease-in-out pb-16 no-print"
      style={{ backgroundImage: `linear-gradient(to bottom, rgba(15, 23, 42, 0.8), rgba(15, 23, 42, 0.92)), url(${theme.bgImage})` }}
    >
      {/* Sticky Premium Header Menu */}
      <header className="sticky top-0 z-50 backdrop-blur-md bg-slate-900/80 border-b border-slate-800 px-4 py-3 shadow-lg no-print">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row items-center justify-between gap-4">
          
          {/* Logo & Subtitle */}
          <div className="flex items-center gap-3">
            <div className={`p-2 rounded-xl bg-${theme.primaryClass}-500/10 border border-${theme.primaryClass}-500/20 text-${theme.primaryClass}-400 shadow-inner`}>
              <Sprout className="w-6 h-6 animate-pulse" />
            </div>
            <div>
              <h1 className="text-white font-semibold text-lg tracking-tight font-hindi">बिहार कृषि विकास एवं युवा उद्यमिता पहल</h1>
              <p className="text-slate-400 text-xs flex items-center gap-1">
                <span>प्रस्तावक: रवि कुमार (BOTANY CHATRA)</span>
                <span className="text-slate-600">•</span>
                <Clock className="w-3.5 h-3.5 text-slate-500" />
                <span>दिनांक: 02 जुलाई 2026</span>
              </p>
            </div>
          </div>

          {/* Interactive Utility Suite */}
          <div className="flex flex-wrap items-center gap-3">
            
            {/* Theme Selector */}
            <div className="bg-slate-950/60 p-1 rounded-lg border border-slate-800 flex items-center gap-1 text-xs">
              <span className="text-slate-500 px-2 font-hindi text-[11px]">पृष्ठभूमि:</span>
              {THEMES.map(t => (
                <button
                  key={t.id}
                  onClick={() => setSelectedThemeId(t.id)}
                  className={`px-3 py-1.5 rounded-md transition-all font-hindi font-medium ${
                    selectedThemeId === t.id 
                      ? `bg-${theme.primaryClass}-500/20 text-${theme.primaryClass}-300 border border-${theme.primaryClass}-500/30` 
                      : 'text-slate-400 hover:text-white'
                  }`}
                >
                  {t.id === 'green' ? 'पॉडी ग्रीन' : 'हार्वेस्ट गोल्ड'}
                </button>
              ))}
            </div>

            {/* Print/Copy Suite */}
            <div className="flex items-center gap-2">
              <button
                onClick={handleSpeak}
                className={`px-3 py-2 rounded-lg text-xs font-medium font-hindi flex items-center gap-1.5 transition-all cursor-pointer ${
                  isSpeaking 
                    ? 'bg-rose-500/20 text-rose-300 border border-rose-500/30 animate-pulse' 
                    : 'bg-slate-800 text-slate-200 border border-slate-700 hover:bg-slate-700 hover:text-white'
                }`}
                title="Hindi TTS Audio Reader"
              >
                {isSpeaking ? <VolumeX className="w-4 h-4 text-rose-400" /> : <Volume2 className="w-4 h-4" />}
                <span>{isSpeaking ? 'वाचन बंद करें' : 'पत्र सुनें'}</span>
              </button>

              <button
                onClick={handleCopyText}
                className={`px-3 py-2 rounded-lg text-xs font-medium font-hindi flex items-center gap-1.5 transition-all cursor-pointer border ${
                  copied 
                    ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/30' 
                    : 'bg-slate-800 text-slate-200 border-slate-700 hover:bg-slate-700 hover:text-white'
                }`}
              >
                {copied ? <Check className="w-4 h-4 text-emerald-400" /> : <Copy className="w-4 h-4" />}
                <span>{copied ? 'कॉपी हो गया' : 'पत्र कॉपी करें'}</span>
              </button>

              <button
                onClick={handlePrint}
                className="px-3 py-2 rounded-lg text-xs font-medium font-hindi flex items-center gap-1.5 bg-gradient-to-r from-emerald-600 to-teal-600 text-white hover:from-emerald-500 hover:to-teal-500 border border-emerald-500/20 shadow-md hover:shadow-emerald-900/30 transition-all cursor-pointer"
                title="प्रिंट या PDF के रूप में सहेजें"
              >
                <Printer className="w-4 h-4" />
                <span>प्रिंट / PDF निकालें</span>
              </button>

              <button
                onClick={handleDownloadTxt}
                className="px-3 py-2 rounded-lg text-xs font-medium font-hindi flex items-center gap-1.5 bg-slate-800 text-slate-200 border border-slate-700 hover:bg-slate-700 hover:text-white transition-all cursor-pointer"
                title="प्लेन टेक्स्ट (.txt) फाइल डाउनलोड करें"
              >
                <Download className="w-4 h-4 text-emerald-400" />
                <span>डाउनलोड (TXT)</span>
              </button>

              <button
                onClick={handleDownloadHtml}
                className="px-3 py-2 rounded-lg text-xs font-medium font-hindi flex items-center gap-1.5 bg-slate-800 text-slate-200 border border-slate-700 hover:bg-slate-700 hover:text-white transition-all cursor-pointer"
                title="सुंदर ऑफलाइन HTML (.html) फाइल डाउनलोड करें"
              >
                <DownloadCloud className="w-4 h-4 text-amber-400" />
                <span>डाउनलोड (HTML)</span>
              </button>

              <button
                onClick={handleDownloadPpt}
                className="px-3 py-2 rounded-lg text-xs font-medium font-hindi flex items-center gap-1.5 bg-slate-800 text-slate-200 border border-slate-700 hover:bg-slate-700 hover:text-white transition-all cursor-pointer"
                title="पावरपॉइंट (.pptx) प्रेजेंटेशन फ़ाइल डाउनलोड करें"
              >
                <Download className="w-4 h-4 text-orange-400" />
                <span>डाउनलोड (PPTX)</span>
              </button>

              <button
                onClick={() => {
                  setCurrentSlideIndex(0);
                  setShowPresenterMode(true);
                }}
                className="px-3 py-2 rounded-lg text-xs font-medium font-hindi flex items-center gap-1.5 bg-slate-800 text-slate-200 border border-slate-700 hover:bg-slate-700 hover:text-white transition-all cursor-pointer"
                title="इंटरैक्टिव पीपीटी प्रेजेंटेशन स्लाइड शो खोलें"
              >
                <Sparkles className="w-4 h-4 text-amber-400 animate-pulse" />
                <span>स्लाइड शो (PPT)</span>
              </button>
            </div>

          </div>
        </div>
      </header>

      {/* Main Split Layout container */}
      <main className="max-w-7xl mx-auto px-4 mt-8 grid grid-cols-1 lg:grid-cols-12 gap-8">
        
        {/* Left Side: Elegant Petition Letter Document Card (70% width) */}
        <section className="lg:col-span-8 flex flex-col gap-4">
          
          {/* Authentic Document Board */}
          <div 
            id="letter-document-root" 
            className="bg-[#fcfbf7] text-slate-800 rounded-3xl shadow-2xl overflow-hidden border-2 border-[#eae4d3] relative print-container"
          >
            {/* Traditional Top Accent Line representing Official Petition */}
            <div className={`h-3 bg-gradient-to-r ${selectedThemeId === 'green' ? 'from-emerald-700 via-emerald-600 to-teal-600' : 'from-amber-700 via-amber-600 to-yellow-600'}`} />

            {/* Document Margin padding */}
            <div className="p-6 md:p-12 font-hindi leading-relaxed print-letter-bg">
              
              {/* Official Seal Header Layout */}
              <div className="flex flex-col items-center text-center border-b-2 border-dashed border-[#dfd8bf] pb-8 mb-8 no-print">
                <div className={`w-20 h-20 rounded-full border-4 border-double ${selectedThemeId === 'green' ? 'border-emerald-600/30' : 'border-amber-600/30'} flex items-center justify-center bg-white shadow-md mb-4`}>
                  <div className={`w-14 h-14 rounded-full ${selectedThemeId === 'green' ? 'bg-emerald-50 text-emerald-700' : 'bg-amber-50 text-amber-700'} flex items-center justify-center`}>
                    <Sprout className="w-8 h-8 stroke-[1.5]" />
                  </div>
                </div>
                <h2 className="text-xl md:text-2xl font-bold tracking-wide text-slate-900 font-yatra uppercase">बिहार कृषि संवर्धन एवं युवा उद्यमिता प्रस्ताव</h2>
                <p className="text-xs text-slate-500 mt-1 max-w-md font-sans uppercase">Aspiration Letter for Agricultural Modernization and Farmers Prosperity</p>
                <div className="flex items-center gap-2 mt-3">
                  <span className="h-[1px] w-12 bg-slate-300"></span>
                  <span className="text-[11px] px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 font-medium font-sans">REG NO: BR-AGRI-2026-0702</span>
                  <span className="h-[1px] w-12 bg-slate-300"></span>
                </div>
              </div>

              {/* Letter Addressee Section */}
              <div className="mb-6 space-y-1 text-slate-800 font-serif text-[15px] md:text-[16px] leading-loose">
                <p className="font-bold">सेवा में,</p>
                <p className="font-bold pl-4">माननीय कृषि मंत्री महोदय,</p>
                <p className="font-bold pl-4">बिहार सरकार, पटना।</p>
              </div>

              {/* Subject (विषय) Highlight Card */}
              <div className={`my-6 p-4 rounded-xl border ${selectedThemeId === 'green' ? 'bg-emerald-50/50 border-emerald-100' : 'bg-amber-50/50 border-amber-100'} shadow-sm`}>
                <p className="font-bold text-slate-900 flex items-start gap-2 text-[15px] md:text-[16px]">
                  <span className={`px-2 py-0.5 rounded text-xs font-sans ${selectedThemeId === 'green' ? 'bg-emerald-600 text-white' : 'bg-amber-600 text-white'} shrink-0 mt-1`}>विषय</span>
                  <span>बिहार में कृषि, किसानों की आय एवं युवा कृषि-उद्यमिता को बढ़ावा देने हेतु सुझाव एवं मुलाकात का अनुरोध।</span>
                </p>
              </div>

              {/* Salutation */}
              <div className="mb-4">
                <p className="font-bold text-slate-900">महोदय,</p>
                <p className="pl-4 text-slate-900">सादर प्रणाम।</p>
              </div>

              {/* Body Paragraph 1 (Ravi's Introduction) */}
              <div className="mb-6 text-justify text-slate-800 font-serif text-[15px] md:text-[16.5px] leading-loose space-y-4">
                <p>
                  मेरा नाम <span className="font-bold text-slate-900 underline decoration-slate-300 decoration-2 underline-offset-4">रवि कुमार</span> है। मैं हिंदुनी, आलमपुर गोनपुरा, फुलवारी शरीफ, पटना, बिहार का निवासी हूँ। मेरी आयु <span className="font-semibold text-slate-900">17 वर्ष 11 माह</span> है। मैंने विज्ञान विषय से 12वीं कक्षा उत्तीर्ण की है और वर्तमान में पाटलिपुत्र विश्वविद्यालय से संबद्ध <span className="font-bold text-slate-900">बी.एस. कॉलेज, दानापुर</span> में <span className="font-bold text-slate-900 bg-emerald-50 px-1 border-b-2 border-emerald-300 py-0.5 rounded inline-flex items-center gap-1">
                    <BookOpen className="w-4 h-4 text-emerald-600 inline no-print" />
                    बी.एससी. (वनस्पति विज्ञान)
                  </span> की पढ़ाई कर रहा हूँ। इस कारण मुझे पौधों, कृषि और प्राकृतिक संसाधनों के बारे में प्रारंभिक वैज्ञानिक जानकारी प्राप्त हो रही है।
                </p>

                {/* Paragraph 2: Agri background & problems */}
                <p>
                  मैं एक किसान परिवार से हूँ। मैंने अपने घर और गाँव में देखा है कि कई परिवारों ने खेती छोड़ दी है या केवल जीवन-यापन के लिए खेती कर रहे हैं। किसान कड़ी मेहनत करते हैं, लेकिन उन्हें उनकी मेहनत के अनुसार पर्याप्त लाभ नहीं मिल पाता। बढ़ती लागत, उचित बाजार की कमी, भंडारण सुविधाओं का अभाव और फसलों के उचित मूल्य न मिलने के कारण खेती युवाओं के लिए आकर्षक नहीं रह गई है।
                </p>

                {/* Paragraph 3: Vision of youth connection */}
                <p>
                  इसी वजह से आज अधिकांश युवा खेती से दूर होकर केवल सरकारी नौकरियों को ही सुरक्षित भविष्य मान रहे हैं। मेरा मानना है कि यदि कृषि को आधुनिक तकनीक, शिक्षा, प्रोसेसिंग उद्योग, डिजिटल मार्केटिंग और उद्यमिता से जोड़ा जाए, तो कृषि फिर से युवाओं के लिए सम्मानजनक और लाभदायक क्षेत्र बन सकती है तथा बिहार की अर्थव्यवस्था को नई दिशा मिल सकती है।
                </p>
              </div>

              {/* Core Suggestion Section Header */}
              <div className="border-t-2 border-[#dfd8bf] pt-6 mt-8 mb-4">
                <p className="font-bold text-[16px] md:text-[18px] text-slate-900 mb-6 flex items-center gap-2">
                  <Sparkles className={`w-5 h-5 ${selectedThemeId === 'green' ? 'text-emerald-600' : 'text-amber-600'}`} />
                  <span>मेरे कुछ प्रमुख सुझाव निम्नलिखित हैं:</span>
                </p>
              </div>

              {/* Suggestions Grid & Highlights */}
              <div className="space-y-4 mb-8">
                {SUGGESTIONS.map((s) => {
                  const isHighlighted = activeSuggestion === s.id;
                  return (
                    <motion.div
                      key={s.id}
                      onClick={() => setActiveSuggestion(s.id === activeSuggestion ? null : s.id)}
                      className={`p-4 md:p-5 rounded-2xl border transition-all duration-300 cursor-pointer relative group ${
                        isHighlighted
                          ? `bg-gradient-to-br ${selectedThemeId === 'green' ? 'from-emerald-50 to-white border-emerald-300 shadow-md shadow-emerald-50' : 'from-amber-50 to-white border-amber-300 shadow-md shadow-amber-50'}`
                          : 'bg-white/60 hover:bg-white border-slate-200/80 hover:border-slate-300 shadow-sm'
                      }`}
                      whileHover={{ scale: 1.01 }}
                    >
                      {/* Suggestion Indicator Ribbon */}
                      <div className="flex items-start gap-4">
                        <div className={`w-10 h-10 rounded-xl flex items-center justify-center shrink-0 shadow-inner transition-all ${
                          isHighlighted 
                            ? (selectedThemeId === 'green' ? 'bg-emerald-600 text-white' : 'bg-amber-600 text-white') 
                            : 'bg-slate-100 text-slate-500 group-hover:bg-slate-200'
                        }`}>
                          <span className="font-bold text-[15px]">{s.id}</span>
                        </div>

                        <div className="space-y-1.5 flex-1">
                          <div className="flex items-center justify-between gap-2 flex-wrap">
                            <h4 className={`font-bold text-[15px] md:text-[16px] tracking-tight transition-colors ${
                              isHighlighted 
                                ? (selectedThemeId === 'green' ? 'text-emerald-900' : 'text-amber-900') 
                                : 'text-slate-900'
                            }`}>
                              {s.title}
                            </h4>
                            <span className={`text-[10px] uppercase font-sans tracking-wider px-2 py-0.5 rounded-full border ${
                              isHighlighted 
                                ? (selectedThemeId === 'green' ? 'bg-emerald-100 text-emerald-800 border-emerald-200' : 'bg-amber-100 text-amber-800 border-amber-200') 
                                : 'bg-slate-100 text-slate-500 border-slate-200'
                            }`}>
                              {s.tag}
                            </span>
                          </div>
                          
                          <p className="text-slate-700 text-[14px] md:text-[15px] font-serif leading-relaxed">
                            {s.originalText}
                          </p>

                          {/* Quick Interactive Tooltip guide */}
                          <div className="pt-2 flex items-center gap-1.5 text-[11px] text-slate-400 font-sans">
                            <span className={`w-1.5 h-1.5 rounded-full ${isHighlighted ? (selectedThemeId === 'green' ? 'bg-emerald-500' : 'bg-amber-500') : 'bg-slate-300'}`}></span>
                            <span>{isHighlighted ? 'समीक्षा बंद करने के लिए क्लिक करें' : 'वैज्ञानिक विश्लेषण देखने के लिए क्लिक करें'}</span>
                          </div>
                        </div>
                      </div>
                    </motion.div>
                  );
                })}
              </div>

              {/* Suggestions Closure */}
              <div className="mb-6 text-justify text-slate-800 font-serif text-[15px] md:text-[16.5px] leading-loose space-y-4">
                <p>
                  मेरा विश्वास है कि यदि सरकार, किसान और युवा मिलकर कार्य करें, तो कृषि बिहार की GDP, रोजगार और समग्र विकास का सबसे मजबूत आधार बन सकती है। मेरा उद्देश्य बिहार के किसानों और युवाओं के लिए सकारात्मक बदलाव लाने में अपना योगदान देना है।
                </p>
                <p>
                  अतः आपसे विनम्र अनुरोध है कि मुझे अपने विचारों और सुझावों पर चर्चा करने का अवसर प्रदान करने की कृपा करें। आपके मार्गदर्शन और सहयोग से बिहार में कृषि और कृषि-उद्यमिता को नई दिशा मिल सकती है।
                </p>
              </div>

              {/* Letter Footer: Sincerely / Sign-off */}
              <div className="mt-12 pt-8 border-t-2 border-dashed border-[#dfd8bf] flex flex-col md:flex-row items-start justify-between gap-8 font-serif">
                
                {/* Left block (Date & Salutation) */}
                <div className="space-y-1.5">
                  <p className="text-slate-600 font-semibold flex items-center gap-1">
                    <Calendar className="w-4 h-4 text-slate-400" />
                    <span>दिनांक: 02 जुलाई 2026</span>
                  </p>
                  <p className="text-slate-500 text-sm flex items-center gap-1">
                    <MapPin className="w-4 h-4 text-slate-400" />
                    <span>स्थान: पटना (बिहार)</span>
                  </p>
                </div>

                {/* Right block (Sender Signature & Details) */}
                <div className="text-slate-900 border-l-2 border-emerald-500 pl-4 space-y-1 md:text-right md:border-l-0 md:border-r-2 md:border-emerald-500 md:pr-4 md:pl-0">
                  <p className="font-bold text-slate-600 uppercase text-xs font-sans tracking-widest">सादर पत्र प्रेषक</p>
                  <p className="font-bold text-lg text-slate-950 font-hindi">रवि कुमार</p>
                  <p className="text-xs text-slate-600">बी.एससी. (वनस्पति विज्ञान) छात्र</p>
                  <p className="text-xs text-slate-500">बी.एस. कॉलेज, दानापुर (पटलिपुत्र विश्वविद्यालय)</p>
                  <p className="text-xs text-slate-500 font-mono">मो: +91 6203023603</p>
                  <p className="text-xs text-slate-500 font-mono">ईमेल: kavikumar11825@gmail.com</p>
                </div>

              </div>

            </div>
          </div>

          {/* Authentic Bihar Farmer Fact Ribbon */}
          <div className="bg-slate-900/60 backdrop-blur-sm border border-slate-800 rounded-2xl p-5 text-slate-300 no-print">
            <h4 className="text-white font-bold font-hindi text-[14px] flex items-center gap-1.5 mb-2">
              <Award className="w-5 h-5 text-emerald-400" />
              <span>कृषि प्रधान बिहार: एक ऐतिहासिक दृष्टिकोण</span>
            </h4>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs font-hindi leading-relaxed mt-3 pt-3 border-t border-slate-800">
              <div className="space-y-1">
                <span className="text-emerald-400 font-bold block">76% आजीविका</span>
                <p className="text-slate-400 text-[11px]">{BIHAR_AGRI_HIGHLIGHTS.gdpShare}</p>
              </div>
              <div className="space-y-1 border-y md:border-y-0 md:border-x border-slate-800 py-3 md:py-0 md:px-4">
                <span className="text-emerald-400 font-bold block">गंगा की जलोढ़ मृदा</span>
                <p className="text-slate-400 text-[11px]">{BIHAR_AGRI_HIGHLIGHTS.gangeticSoil}</p>
              </div>
              <div className="space-y-1">
                <span className="text-emerald-400 font-bold block">प्रचुर जल संसाधन</span>
                <p className="text-slate-400 text-[11px]">{BIHAR_AGRI_HIGHLIGHTS.waterAbundance}</p>
              </div>
            </div>
          </div>

        </section>

        {/* Right Side: Interactive Companion Panel (30% width) */}
        <section className="lg:col-span-4 flex flex-col gap-6 no-print">
          
          {/* 1. Ravi Kumar's Student Botany Profile ID */}
          <div className="bg-slate-900/80 backdrop-blur-md border border-slate-800 rounded-3xl overflow-hidden shadow-xl">
            {/* ID Card Top Banner */}
            <div className={`p-4 bg-gradient-to-r ${selectedThemeId === 'green' ? 'from-emerald-950 to-emerald-900' : 'from-amber-950 to-amber-900'} border-b border-slate-800 flex items-center justify-between`}>
              <div className="flex items-center gap-2">
                <BookOpen className={`w-4 h-4 ${selectedThemeId === 'green' ? 'text-emerald-400' : 'text-amber-400'}`} />
                <span className="text-slate-300 font-sans text-[11px] font-semibold tracking-wider uppercase">PATLIPUTRA UNIVERSITY</span>
              </div>
              <span className={`text-[9px] px-2 py-0.5 rounded-full font-bold bg-${theme.primaryClass}-500/15 text-${theme.primaryClass}-400 border border-${theme.primaryClass}-500/30 font-sans`}>
                B.SC STUDENT
              </span>
            </div>

            {/* Profile Content */}
            <div className="p-6 flex flex-col items-center text-center">
              
              {/* Avatar representation with Botanical decoration */}
              <div className="relative mb-4">
                <div className={`w-20 h-20 rounded-2xl bg-gradient-to-tr ${selectedThemeId === 'green' ? 'from-emerald-600 to-teal-500' : 'from-amber-600 to-yellow-500'} p-0.5 shadow-md`}>
                  <div className="w-full h-full bg-slate-950 rounded-[14px] flex items-center justify-center text-white text-2xl font-bold font-hindi">
                    रवि
                  </div>
                </div>
                <div className={`absolute -bottom-1 -right-1 p-1.5 rounded-full bg-slate-900 border border-slate-800 text-${theme.primaryClass}-400 shadow-md`}>
                  <Sprout className="w-4 h-4" />
                </div>
              </div>

              {/* Personal Details */}
              <h3 className="text-white font-bold text-lg font-hindi">रवि कुमार</h3>
              <p className="text-slate-400 text-xs mt-1 font-hindi">बी.एससी. (वनस्पति विज्ञान) छात्र, बी.एस. कॉलेज, दानापुर</p>
              
              {/* Quick Bio Statistics */}
              <div className="grid grid-cols-2 gap-2 w-full mt-4 py-3 border-y border-slate-800/80 text-xs">
                <div className="text-center">
                  <span className="text-slate-500 block">आयु</span>
                  <span className="text-slate-200 font-medium font-hindi">17 वर्ष 11 माह</span>
                </div>
                <div className="text-center border-l border-slate-800">
                  <span className="text-slate-500 block">विषय समूह</span>
                  <span className="text-slate-200 font-medium font-hindi">विज्ञान (विज्ञान 12वीं)</span>
                </div>
              </div>

              {/* Direct Communication Channels */}
              <div className="w-full mt-4 space-y-2.5">
                
                <a 
                  href="tel:6203023603" 
                  className="w-full p-2.5 rounded-xl bg-slate-950 border border-slate-800/80 flex items-center justify-between hover:border-slate-700 hover:bg-slate-900 text-left transition-all group"
                >
                  <div className="flex items-center gap-2">
                    <Phone className={`w-4 h-4 text-${theme.primaryClass}-400`} />
                    <span className="text-xs text-slate-400 font-mono group-hover:text-white transition-colors">+91 6203023603</span>
                  </div>
                  <ChevronRight className="w-3.5 h-3.5 text-slate-500" />
                </a>

                <a 
                  href="mailto:kavikumar11825@gmail.com" 
                  className="w-full p-2.5 rounded-xl bg-slate-950 border border-slate-800/80 flex items-center justify-between hover:border-slate-700 hover:bg-slate-900 text-left transition-all group"
                >
                  <div className="flex items-center gap-2">
                    <Mail className={`w-4 h-4 text-${theme.primaryClass}-400`} />
                    <span className="text-xs text-slate-400 font-mono group-hover:text-white transition-colors text-ellipsis overflow-hidden max-w-[170px]">kavikumar11825@gmail.com</span>
                  </div>
                  <ChevronRight className="w-3.5 h-3.5 text-slate-500" />
                </a>

                <div className="p-2.5 rounded-xl bg-slate-950/60 border border-slate-800/50 flex items-start gap-2 text-left">
                  <MapPin className="w-4 h-4 text-slate-500 mt-0.5 shrink-0" />
                  <span className="text-[11px] text-slate-400 font-hindi leading-relaxed">हिंदुनी, आलमपुर गोनपुरा, फुलवारी शरीफ, पटना, बिहार</span>
                </div>

              </div>

            </div>
          </div>

          {/* 2. Interactive Selected Suggestion Spotlight */}
          <AnimatePresence mode="wait">
            {activeSuggestion !== null ? (
              <motion.div
                key={`spotlight-${activeSuggestion}`}
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className={`bg-slate-900/90 border-2 ${selectedThemeId === 'green' ? 'border-emerald-800/60' : 'border-amber-800/60'} rounded-3xl p-6 shadow-xl relative overflow-hidden`}
              >
                {/* Decorative glow bg */}
                <div className={`absolute -right-10 -top-10 w-32 h-32 rounded-full ${selectedThemeId === 'green' ? 'bg-emerald-500/10' : 'bg-amber-500/10'} blur-2xl`} />

                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-2">
                    <div className={`p-2 rounded-xl bg-${theme.primaryClass}-500/10 border border-${theme.primaryClass}-500/20 text-${theme.primaryClass}-400`}>
                      {renderIcon(SUGGESTIONS[activeSuggestion - 1].iconName)}
                    </div>
                    <span className="text-slate-400 font-sans text-xs font-bold uppercase tracking-wider">सुझाव विश्लेषण #{activeSuggestion}</span>
                  </div>
                  <button 
                    onClick={() => setActiveSuggestion(null)}
                    className="text-slate-500 hover:text-white hover:bg-slate-800 p-1 rounded-full transition-colors"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>

                <h3 className="text-white font-bold text-[16px] font-hindi mb-2">
                  {SUGGESTIONS[activeSuggestion - 1].title}
                </h3>

                <p className="text-slate-300 text-[13px] font-hindi leading-relaxed bg-slate-950/40 p-3 rounded-xl border border-slate-800/50 mb-4 font-serif">
                  &ldquo;{SUGGESTIONS[activeSuggestion - 1].originalText}&rdquo;
                </p>

                <div className="space-y-3 font-hindi">
                  <h4 className={`text-xs uppercase tracking-wider text-${theme.primaryClass}-400 font-semibold flex items-center gap-1`}>
                    <Sprout className="w-3.5 h-3.5" />
                    <span>वैज्ञानिक एवं नीतिगत विश्लेषण</span>
                  </h4>
                  <p className="text-slate-400 text-xs leading-relaxed font-serif">
                    {SUGGESTIONS[activeSuggestion - 1].detailedContext}
                  </p>

                  {SUGGESTIONS[activeSuggestion - 1].stats && (
                    <div className="mt-4 pt-3 border-t border-slate-800/80 flex items-center gap-2 text-xs text-slate-300 bg-slate-950/50 p-2.5 rounded-lg border border-slate-800/40">
                      <Award className={`w-4 h-4 text-${theme.primaryClass}-400 shrink-0`} />
                      <span className="font-serif font-medium">{SUGGESTIONS[activeSuggestion - 1].stats}</span>
                    </div>
                  )}
                </div>

              </motion.div>
            ) : (
              <div className="bg-slate-900/50 border border-slate-800 border-dashed rounded-3xl p-6 text-center shadow-md">
                <FileText className="w-8 h-8 text-slate-600 mx-auto mb-3" />
                <h4 className="text-slate-300 font-bold text-xs uppercase tracking-wider font-sans">सुझाव विश्लेषण केंद्र</h4>
                <p className="text-slate-500 text-xs mt-1.5 font-hindi leading-relaxed">
                  पत्र में लिखे गए किसी भी <span className={`text-${theme.primaryClass}-400 font-bold`}>सुझाव</span> पर क्लिक करें और रवि के विस्तृत वैज्ञानिक एवं जमीनी दृष्टिकोण को यहाँ जानें।
                </p>
              </div>
            )}
          </AnimatePresence>

          {/* 3. Ravi's Botany student Crop lab */}
          <div className="bg-slate-900/80 backdrop-blur-md border border-slate-800 rounded-3xl p-6 shadow-xl">
            <div className="flex items-center gap-2 mb-4">
              <Sprout className="w-5 h-5 text-emerald-400" />
              <h3 className="text-white font-bold font-hindi text-[14px]">रवि की वनस्पति प्रयोगशाला (Bihar Crops Lab)</h3>
            </div>

            <p className="text-slate-400 text-xs font-hindi leading-relaxed mb-4">
              रवि कुमार दानापुर के बी.एस. कॉलेज में वनस्पति विज्ञान (Botany) के छात्र हैं। उन्होंने बिहार के इन प्रमुख उत्पादों के वैज्ञानिक मूल्यवर्धन (Scientific Value-Addition) का प्रस्ताव दिया है:
            </p>

            {/* Crop Selectors tabs */}
            <div className="grid grid-cols-4 gap-1.5 mb-4 p-1 bg-slate-950 rounded-xl border border-slate-800">
              {BOTANICAL_CROPS.map((crop, index) => (
                <button
                  key={index}
                  onClick={() => setActiveCropIndex(index)}
                  className={`py-2 px-1 rounded-lg text-center transition-all text-[11px] font-hindi font-semibold ${
                    activeCropIndex === index 
                      ? `bg-${theme.primaryClass}-500/10 text-${theme.primaryClass}-300 border border-${theme.primaryClass}-500/30` 
                      : 'text-slate-400 hover:text-white hover:bg-slate-900'
                  }`}
                >
                  {crop.hindiName.split(' ')[0]}
                </button>
              ))}
            </div>

            {/* Crop details show */}
            <AnimatePresence mode="wait">
              <motion.div
                key={activeCropIndex}
                initial={{ opacity: 0, x: 10 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: -10 }}
                className="bg-slate-950/50 border border-slate-800 rounded-2xl p-4 font-hindi space-y-3"
              >
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <h4 className="text-white font-bold text-sm">{BOTANICAL_CROPS[activeCropIndex].hindiName}</h4>
                    <span className="text-emerald-400 text-[11px] italic font-sans block">{BOTANICAL_CROPS[activeCropIndex].scientificName}</span>
                  </div>
                  {renderIcon(BOTANICAL_CROPS[activeCropIndex].iconName, "w-8 h-8 opacity-90")}
                </div>

                <div className="space-y-1.5 text-xs border-t border-slate-800/80 pt-2.5">
                  <p className="flex justify-between gap-2 text-[11px]">
                    <span className="text-slate-500">कुल / Family:</span>
                    <span className="text-slate-300 font-medium">{BOTANICAL_CROPS[activeCropIndex].family}</span>
                  </p>
                  <p className="flex justify-between gap-2 text-[11px]">
                    <span className="text-slate-500">फसल सीजन:</span>
                    <span className="text-slate-300 font-medium">{BOTANICAL_CROPS[activeCropIndex].harvestSeason}</span>
                  </p>
                </div>

                <p className="text-slate-400 text-[11.5px] leading-relaxed font-serif pt-1">
                  {BOTANICAL_CROPS[activeCropIndex].description}
                </p>

                <div className="bg-emerald-950/20 border border-emerald-900/30 p-2.5 rounded-xl flex items-start gap-2">
                  <Award className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
                  <p className="text-[10.5px] text-slate-300 leading-relaxed font-serif">
                    <span className="font-bold text-emerald-300 block">बिहार की ताकत:</span>
                    {BOTANICAL_CROPS[activeCropIndex].biharStats}
                  </p>
                </div>

              </motion.div>
            </AnimatePresence>

          </div>

          {/* 4. Interactive Minister's Desk Simulation */}
          <div className="bg-slate-900/80 backdrop-blur-md border border-slate-800 rounded-3xl p-6 shadow-xl relative overflow-hidden">
            <div className="absolute -left-10 -bottom-10 w-32 h-32 rounded-full bg-emerald-500/5 blur-2xl" />

            <div className="flex items-center gap-2 mb-3 border-b border-slate-800/80 pb-3">
              <div className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse" />
              <h3 className="text-white font-bold font-hindi text-[14px]">माननीय कृषि मंत्री का डेस्क (सिमुलेशन)</h3>
            </div>

            {ministerDecision === 'pending' ? (
              <div className="space-y-4 font-hindi">
                <p className="text-slate-400 text-xs leading-relaxed">
                  आप इस समय बिहार सरकार के <strong>कृषि मंत्री</strong> के रूप में कार्य कर रहे हैं। रवि कुमार के इस प्रस्ताव पत्र को पढ़ने के बाद आप क्या निर्णय लेंगे?
                </p>

                <div className="space-y-2.5">
                  <div className="space-y-1">
                    <label className="text-slate-500 text-[10px] uppercase font-sans tracking-wider block">मुलाकात की तिथि तय करें</label>
                    <div className="grid grid-cols-2 gap-2">
                      <input 
                        type="date" 
                        value={meetingDate}
                        onChange={(e) => setMeetingDate(e.target.value)}
                        className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-emerald-500 font-mono"
                      />
                      <input 
                        type="time" 
                        value={meetingTime}
                        onChange={(e) => setMeetingTime(e.target.value)}
                        className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-emerald-500 font-mono"
                      />
                    </div>
                  </div>

                  <div className="space-y-1">
                    <label className="text-slate-500 text-[10px] uppercase font-sans tracking-wider block">रवि के लिए आधिकारिक संदेश / टिप्पणी</label>
                    <textarea 
                      value={ministerNote}
                      onChange={(e) => setMinisterNote(e.target.value)}
                      rows={3}
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-300 focus:outline-none focus:border-emerald-500 font-serif leading-relaxed"
                      placeholder="मंत्री जी की टिप्पणी..."
                    />
                  </div>
                </div>

                <div className="flex gap-2 pt-1">
                  <button
                    onClick={() => {
                      setMinisterDecision('approved');
                      setShowApprovalNotice(true);
                    }}
                    className="flex-1 py-2.5 px-3 bg-gradient-to-r from-emerald-600 to-teal-600 text-white font-bold text-xs rounded-xl shadow-lg hover:from-emerald-500 hover:to-teal-500 transition-all text-center cursor-pointer"
                  >
                    सुझाव स्वीकृत करें &amp; पत्र भेजें
                  </button>
                </div>
              </div>
            ) : (
              <motion.div 
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="space-y-4 font-hindi text-slate-300"
              >
                <div className="bg-emerald-950/30 border border-emerald-900/50 p-3 rounded-2xl flex items-start gap-3">
                  <CheckCircle className="w-6 h-6 text-emerald-400 shrink-0 mt-0.5" />
                  <div>
                    <h4 className="text-white font-bold text-xs">मुलाकात निर्धारित एवं प्रस्ताव स्वीकृत!</h4>
                    <p className="text-[11px] text-slate-400 mt-1 font-serif">
                      रवि कुमार को आधिकारिक सरकारी पत्र के माध्यम से सूचित कर दिया गया है।
                    </p>
                  </div>
                </div>

                {/* Simulated Official Response Letter */}
                <div className="bg-white text-slate-900 p-4 rounded-xl shadow-inner text-[11px] leading-relaxed font-serif relative">
                  <div className="absolute top-2 right-2 text-[9px] uppercase tracking-wider text-slate-400 font-bold font-sans">OFFICIAL COPY</div>
                  
                  {/* Small Seal */}
                  <div className="border-b border-slate-200 pb-2 mb-2 flex items-center gap-1.5">
                    <Sprout className="w-3.5 h-3.5 text-emerald-700" />
                    <span className="font-bold text-[10px] tracking-wide text-emerald-800">कृषि मंत्रालय, बिहार सरकार</span>
                  </div>

                  <p className="text-[10px] font-sans font-bold text-slate-500">पत्रांक: कृ-मंत्रा/२०२६/प्र-१०८</p>
                  
                  <div className="mt-2 space-y-1">
                    <p className="font-bold">प्रिय रवि कुमार,</p>
                    <p className="text-slate-800 leading-loose text-justify text-[10.5px]">
                      {ministerNote}
                    </p>
                  </div>

                  <div className="mt-3 pt-2 border-t border-slate-100 flex justify-between items-end">
                    <div className="text-[9px] text-slate-500 font-sans">
                      <p className="font-bold text-slate-700">तय तिथि: {new Date(meetingDate).toLocaleDateString('hi-IN', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
                      <p className="font-bold text-slate-700">समय: {meetingTime} बजे</p>
                      <p>स्थान: विकास भवन, पटना</p>
                    </div>
                    <div className="text-right">
                      <div className="w-10 h-4 bg-slate-200/50 border border-slate-300 rounded inline-block mb-1 opacity-60"></div>
                      <p className="font-bold text-slate-800 font-hindi text-[9.5px]">कृषि मंत्री</p>
                      <p className="text-[8.5px] text-slate-400">बिहार सरकार</p>
                    </div>
                  </div>
                </div>

                <button
                  onClick={() => {
                    setMinisterDecision('pending');
                    setShowApprovalNotice(false);
                  }}
                  className="w-full py-2 bg-slate-800 text-slate-300 text-xs font-semibold rounded-xl hover:bg-slate-700 hover:text-white transition-all cursor-pointer"
                >
                  पुनः सिमुलेशन शुरू करें
                </button>
              </motion.div>
            )}
          </div>

        </section>

      </main>

      {/* Floating toast notification for copy */}
      <AnimatePresence>
        {copied && (
          <motion.div 
            initial={{ opacity: 0, y: 50 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: 50 }}
            className="fixed bottom-6 left-1/2 transform -translate-x-1/2 z-50 bg-slate-900 border border-slate-800 px-4 py-3 rounded-2xl flex items-center gap-2 shadow-2xl text-white text-xs font-hindi"
          >
            <CheckCircle className="w-4 h-4 text-emerald-400" />
            <span>पूरा पत्र सफलतापूर्वक क्लिपबोर्ड पर कॉपी हो गया है!</span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Virtual modal popup for Minister decision success */}
      <AnimatePresence>
        {showApprovalNotice && (
          <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-slate-900 border border-slate-800 rounded-3xl p-6 max-w-md w-full shadow-2xl text-center font-hindi text-slate-300"
            >
              <div className="w-16 h-16 rounded-full bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 flex items-center justify-center mx-auto mb-4">
                <CheckCircle className="w-8 h-8" />
              </div>
              
              <h3 className="text-white font-bold text-lg">बिहार कृषि संवर्धन पत्र स्वीकृत!</h3>
              
              <p className="text-xs text-slate-400 mt-2 leading-relaxed font-serif">
                माननीय मंत्री जी ने रवि कुमार के सातों प्रस्तावों को ध्यानपूर्वक पढ़कर सैद्धांतिक स्वीकृति दे दी है। जुलाई {new Date(meetingDate).getFullYear()} को विकास भवन, पटना में मुलाकात का समय निर्धारित किया गया है।
              </p>

              <div className="bg-slate-950 p-3 rounded-xl text-xs text-left my-4 border border-slate-800 font-serif leading-relaxed">
                <p className="font-bold text-white mb-1">मुलाकात विवरण:</p>
                <p className="text-slate-400">दिनांक: {new Date(meetingDate).toLocaleDateString('hi-IN', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
                <p className="text-slate-400">समय: {meetingTime} बजे</p>
                <p className="text-slate-400">स्थान: कृषि मंत्रालय कार्यालय, विकास भवन, पटना</p>
              </div>

              <button
                onClick={() => setShowApprovalNotice(false)}
                className="w-full py-2.5 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-bold text-xs rounded-xl cursor-pointer"
              >
                विवरण पत्र देखें
              </button>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Quick guide for users footer */}
      <footer className="max-w-7xl mx-auto px-4 mt-8 text-center text-slate-500 text-xs font-hindi border-t border-slate-800/40 pt-6 no-print">
        <p>© 2026 बिहार कृषि विकास पहल • रवि कुमार, बी.एससी. (वनस्पति विज्ञान) छात्र, दानापुर</p>
        <p className="text-slate-600 text-[10px] mt-1">
          यह एक इंटरैक्टिव प्रस्ताव पत्र और सिमुलेशन डैशबोर्ड है। इसे प्रिंट करने के लिए "प्रिंट / PDF निकालें" पर क्लिक करें।
        </p>
      </footer>

      {/* Hidden layout specifically customized for paper PRINTING (CSS print style overrides) */}
      <div className="hidden print:block print-container font-serif">
        <div className="bg-white text-slate-950 p-8 w-full leading-relaxed text-[15px]">
          
          {/* Header Seal */}
          <div className="text-center border-b-2 border-slate-300 pb-4 mb-6">
            <h1 className="text-2xl font-bold tracking-wide">बिहार कृषि संवर्धन एवं युवा उद्यमिता प्रस्ताव</h1>
            <p className="text-xs uppercase tracking-wider text-slate-600 mt-1">Aspiration Letter for Agricultural Modernization and Farmers Prosperity</p>
            <p className="text-[10px] font-mono text-slate-500 mt-1">REG NO: BR-AGRI-2026-0702</p>
          </div>

          {/* Addressee */}
          <div className="mb-6 space-y-1">
            <p className="font-bold">सेवा में,</p>
            <p className="font-bold pl-4">माननीय कृषि मंत्री महोदय,</p>
            <p className="font-bold pl-4">बिहार सरकार, पटना।</p>
          </div>

          {/* Subject */}
          <div className="border border-slate-300 p-3 bg-slate-50/50 rounded-lg mb-6">
            <p className="font-bold text-slate-900 text-[14px]">
              विषय: बिहार में कृषि, किसानों की आय एवं युवा कृषि-उद्यमिता को बढ़ावा देने हेतु सुझाव एवं मुलाकात का अनुरोध।
            </p>
          </div>

          {/* Salutation */}
          <div className="mb-4">
            <p className="font-bold">महोदय,</p>
            <p className="pl-4">सादर प्रणाम।</p>
          </div>

          {/* Intro */}
          <div className="space-y-4 text-justify leading-loose text-[14px]">
            <p>
              मेरा नाम <strong>रवि कुमार</strong> है। मैं हिंदुनी, आलमपुर गोनपुरा, फुलवारी शरीफ, पटना, बिहार का निवासी हूँ। मेरी आयु 17 वर्ष 11 माह है। मैंने विज्ञान विषय से 12वीं कक्षा उत्तीर्ण की है और वर्तमान में पाटलिपुत्र विश्वविद्यालय से संबद्ध बी.एस. कॉलेज, दानापुर में बी.एससी. (वनस्पति विज्ञान) की पढ़ाई कर रहा हूँ। इस कारण मुझे पौधों, कृषि और प्राकृतिक संसाधनों के बारे में प्रारंभिक वैज्ञानिक जानकारी प्राप्त हो रही है।
            </p>
            <p>
              मैं एक किसान परिवार से हूँ। मैंने अपने घर और गाँव में देखा है कि कई परिवारों ने खेती छोड़ दी है या केवल जीवन-यापन के लिए खेती कर रहे हैं। किसान कड़ी मेहनत करते हैं, लेकिन उन्हें उनकी मेहनत के अनुसार पर्याप्त लाभ नहीं मिल पाता। बढ़ती लागत, उचित बाजार की कमी, भंडारण सुविधाओं का अभाव और फसलों के उचित मूल्य न मिलने के कारण खेती युवाओं के लिए आकर्षक नहीं रह गई है।
            </p>
            <p>
              इसी वजह से आज अधिकांश युवा खेती से दूर होकर केवल सरकारी नौकरियों को ही सुरक्षित भविष्य मान रहे हैं। मेरा मानना है कि यदि कृषि को आधुनिक तकनीक, शिक्षा, प्रोसेसिंग उद्योग, डिजिटल मार्केटिंग और उद्यमिता से जोड़ा जाए, तो कृषि फिर से युवाओं के लिए सम्मानजनक और लाभदायक क्षेत्र बन सकती है तथा बिहार की अर्थव्यवस्था को नई दिशा मिल सकती है।
            </p>
          </div>

          <p className="font-bold text-[15px] my-4 pt-4 border-t border-slate-200">मेरे कुछ सुझाव निम्नलिखित हैं:</p>

          {/* Core 7 suggestions */}
          <div className="space-y-3.5 pl-2 leading-relaxed text-[13.5px]">
            <p><strong>1.</strong> प्रत्येक जिले में <strong>Young Agri Entrepreneur Clubs</strong> की स्थापना की जाए, जहाँ युवाओं को आधुनिक खेती और कृषि-उद्यमिता का प्रशिक्षण दिया जाए.</p>
            <p><strong>2.</strong> युवाओं और किसानों को AI, डिजिटल मार्केटिंग, ड्रोन तकनीक और आधुनिक कृषि पद्धतियों की जानकारी एवं प्रशिक्षण उपलब्ध कराया जाए.</p>
            <p><strong>3.</strong> मखाना, मक्का, लीची, सब्जियों और अन्य स्थानीय उत्पादों की प्रोसेसिंग, पैकेजिंग और ब्रांडिंग को बढ़ावा दिया जाए, ताकि किसानों को अधिक लाभ मिल सके.</p>
            <p><strong>4.</strong> कृषि आधारित स्टार्टअप्स के लिए विशेष अनुदान, मार्गदर्शन और इनक्यूबेशन सुविधाएँ उपलब्ध कराई जाएँ.</p>
            <p><strong>5.</strong> स्कूलों और कॉलेजों में कृषि-उद्यमिता, एग्री-टेक और आधुनिक खेती से संबंधित कार्यक्रम आयोजित किए जाएँ, जिससे युवा कृषि को एक करियर और व्यवसाय के रूप में अपनाने के लिए प्रेरित हों.</p>
            <p><strong>6.</strong> छोटे किसानों के लिए कोल्ड स्टोरेज, वेयरहाउस, बेहतर बाजार व्यवस्था और उचित मूल्य सुनिश्चित करने की दिशा में विशेष योजनाएँ लागू की जाएँ.</p>
            <p><strong>7.</strong> युवाओं को केवल नौकरी खोजने वाला नहीं, बल्कि कृषि के माध्यम से रोजगार देने वाला उद्यमी बनने के लिए प्रोत्साहित किया जाए.</p>
          </div>

          {/* Outro */}
          <div className="space-y-4 text-justify leading-loose text-[14px] mt-6 pt-4 border-t border-slate-200">
            <p>
              मेरा विश्वास है कि यदि सरकार, किसान और युवा मिलकर कार्य करें, तो कृषि बिहार की GDP, रोजगार और समग्र विकास का सबसे मजबूत आधार बन सकती है। मेरा उद्देश्य बिहार के किसानों और युवाओं के लिए सकारात्मक बदलाव लाने में अपना योगदान देना है।
            </p>
            <p>
              अतः आपसे विनम्र अनुरोध है कि मुझे अपने विचारों और सुझावों पर चर्चा करने का अवसर प्रदान करने की कृपा करें। आपके मार्गदर्शन और सहयोग से बिहार में कृषि और कृषि-उद्यमिता को नई दिशा मिल सकती है।
            </p>
          </div>

          {/* Sign-off */}
          <div className="mt-10 pt-6 border-t border-slate-300 flex justify-between items-start text-xs">
            <div>
              <p><strong>दिनांक:</strong> 02 जुलाई 2026</p>
              <p><strong>स्थान:</strong> पटना, बिहार</p>
            </div>
            <div className="text-right space-y-1">
              <p><strong>सधन्यवाद,</strong></p>
              <p className="font-bold text-[14px] mt-2">रवि कुमार</p>
              <p>बी.एससी. (वनस्पति विज्ञान) छात्र, बी.एस. कॉलेज, दानापुर</p>
              <p>पता: हिंदुनी, आलमपुर गोनपुरा, फुलवारी शरीफ, पटना, बिहार</p>
              <p>मोबाइल: 6203023603 | ईमेल: kavikumar11825@gmail.com</p>
            </div>
          </div>

        </div>
      </div>

    </div>
  );
}
