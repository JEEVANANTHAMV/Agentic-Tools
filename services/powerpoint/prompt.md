Act as an expert Presentation Designer and Frontend Developer. Your task is to generate a multi-slide presentation in a single HTML file that is compatible with a **Hybrid HTML-to-PPTX Engine**.

**Design Requirements:**
1.  **Dimensions:** Every slide must be wrapped in a `<div class="ppt-slide">`. This div is strictly **1280px wide and 720px high**.
2.  **Styling:** Use **Tailwind CSS** utility classes exclusively. Assume a standard Tailwind configuration is available. 
3.  **Typography:** * Use `Source Code Pro` for headings and `Roboto Flex` for body text. 
    * Use explicit pixel values for text sizes (e.g., `text-[24px]`) to ensure the Python engine calculates positions accurately.
4.  **Layout:** Use Flexbox and Grid for positioning. Ensure no content overflows the 1280x720 boundary.
5.  **Visuals:** Use high-contrast colors, gradients (`bg-gradient-to-br`), and Material Icons (`<i class="material-icons">icon_name</i>`).
6.  **Structure:** * Return one single HTML block.
    * Include a `<style>` block in the `<head>` defining the `.ppt-slide` class:
        ```css
        .ppt-slide { position: relative; width: 1280px; height: 720px; overflow: hidden; background: white; }
        ```

**Content Task:**
Generate a [Number of Slides] slide deck about [Your Topic]. 
* Slide 1: Title Slide (Title, Subtitle, Company Name).
* Slide 2: Problem/Solution (Two-column grid).
* Slide 3: Key Features (Icon-based grid).
* Slide 4: Contact/Thank You.

---

### Example Output for "InnoSynth" (Based on your sample)

If you use the prompt above, the AI will produce code like this, which your Python service can then process into a `.pptx`:

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Source+Code+Pro:wght@700&family=Roboto+Flex:wght@300;400&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <style>
        body { font-family: "Roboto Flex", sans-serif; margin: 0; padding: 0; background: #333; }
        .ppt-slide { position: relative; width: 1280px; height: 720px; margin: 20px auto; overflow: hidden; background: white; box-shadow: 0 20px 25px -5px rgb(0 0 0 / 0.1); }
        h1, h2 { font-family: "Source Code Pro", monospace; }
        .accent-bar { width: 100px; height: 8px; background-color: #A6C40D; }
    </style>
</head>
<body>

    <div class="ppt-slide flex flex-col justify-center items-center px-20">
        <div class="text-center">
            <h1 class="text-[80px] font-bold text-black leading-tight">InnoSynth</h1>
            <div class="accent-bar mx-auto my-8"></div>
            <h2 class="text-[32px] font-semibold text-gray-800">
                World's 1st Chat-based AI Agent Development Platform
            </h2>
            <p class="text-[24px] text-[#A6C40D] font-bold mt-4 uppercase tracking-widest">Without Coding</p>
        </div>
        <div class="absolute bottom-10 flex items-center text-gray-400">
            <i class="material-icons mr-2">location_on</i>
            <span class="text-[18px]">Tiruppur, India</span>
        </div>
    </div>

    <div class="ppt-slide p-16 bg-slate-50">
        <h2 class="text-[40px] font-bold text-black mb-12 border-l-8 border-[#A6C40D] pl-6">Core Capabilities</h2>
        <div class="grid grid-cols-3 gap-10">
            <div class="bg-white p-8 rounded-xl shadow-sm border border-gray-100">
                <i class="material-icons text-[48px] text-[#A6C40D] mb-4">chat</i>
                <h3 class="text-[24px] font-bold mb-2">Natural Language</h3>
                <p class="text-[18px] text-gray-600">Describe your agent in plain English; we handle the logic.</p>
            </div>
            <div class="bg-white p-8 rounded-xl shadow-sm border border-gray-100">
                <i class="material-icons text-[48px] text-[#A6C40D] mb-4">bolt</i>
                <h3 class="text-[24px] font-bold mb-2">Instant Deploy</h3>
                <p class="text-[18px] text-gray-600">Go from concept to live API in under 60 seconds.</p>
            </div>
            <div class="bg-white p-8 rounded-xl shadow-sm border border-gray-100">
                <i class="material-icons text-[48px] text-[#A6C40D] mb-4">hub</i>
                <h3 class="text-[24px] font-bold mb-2">Multi-Tool</h3>
                <p class="text-[18px] text-gray-600">Connect to 100+ apps via built-in integration nodes.</p>
            </div>
        </div>
    </div>

</body>
</html>
```