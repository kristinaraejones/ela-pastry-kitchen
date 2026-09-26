// Pastry Passport — one treat per week of the term.
// Each finished week (all 5 plates served) unlocks that week's treat.
// Photos: drop a picture at images/treats/<slug>.jpg and fill in photoCredit
// (Wikimedia Commons works well — keep the photographer's name and license).
// Step tags: "kid" = Kid job, "grown" = Grown-up job, "together" = Together.
const BK_TREATS = [
  {
    week: 1, slug: "crema-catalana", city: "Barcelona", country: "Spain",
    name: "Crema Catalana", altName: "Also called crema de Sant Josep",
    tagline: "A Catalan custard with a crackly burnt-sugar top",
    photoCredit: "",
    where: "Catalonia, in northeast Spain", whereSub: "Capital: Barcelona · Languages: Catalan and Spanish",
    story: "Cooks in Catalonia have been making this lemony, cinnamon-scented custard for centuries. Recipes for it show up in Catalan cookbooks from the Middle Ages, which makes it one of the oldest custard desserts in Europe.",
    facts: [
      "Families traditionally make it on <b>March 19, Saint Joseph’s Day</b>, which is Father’s Day in Spain. That’s where the name <i>crema de Sant Josep</i> comes from.",
      "It’s a cousin of French <b>crème brûlée</b>, but it uses milk instead of cream and gets its flavor from lemon peel and cinnamon.",
      "The sugar on top was traditionally burnt with a <b>round iron</b> heated on the stove. Today most cooks use a kitchen torch."
    ],
    recipe: {
      serves: "Serves 4", time: "20 min + 2 hr chilling", tools: "Whisk · saucepan · kitchen torch",
      ingredients: ["500 ml (2 cups) whole milk", "Peel of 1 lemon, in wide strips", "1 cinnamon stick", "4 egg yolks", "100 g (½ cup) sugar, plus extra for the tops", "2 tbsp cornstarch"],
      steps: [
        ["Warm the milk with the lemon peel and cinnamon until it steams. Turn off the heat, let it rest 10 minutes, then fish out the peel and cinnamon.", "together"],
        ["Whisk the yolks, sugar and cornstarch in a bowl until pale and smooth.", "kid"],
        ["Slowly whisk the warm milk into the yolks. Pour it all back into the pot and stir over low heat until thick like pudding, 5–8 minutes.", "together"],
        ["Pour into shallow dishes, let cool, then chill for at least 2 hours.", "kid"],
        ["Sprinkle a thin layer of sugar on top and torch it until it turns amber and glassy.", "grown"],
        ["Tap the top with a spoon. Crack! Now it’s ready to eat.", "kid"]
      ]
    }
  },
  {
    week: 2, slug: "concha", city: "Mexico City", country: "Mexico",
    name: "Conchas", altName: "Pan dulce · sweet bread",
    tagline: "Soft sweet rolls with a crunchy seashell topping",
    photoCredit: "",
    where: "Mexico", whereSub: "Capital: Mexico City · Language: Spanish",
    story: "Conchas are soft, sweet bread rolls with a crunchy sugar topping scored to look like a seashell. <i>Concha</i> means “shell” in Spanish. They’re one of the most loved kinds of <i>pan dulce</i> (sweet bread), a tradition that grew in the 1800s when European-style baking blended with Mexican tastes.",
    facts: [
      "In a Mexican <b>panadería</b> you usually pick your own breads with a tray and tongs, then pay at the counter.",
      "The topping comes in <b>vanilla and chocolate</b>, and sometimes in pink or other colors.",
      "Many families eat pan dulce with <b>hot chocolate or café de olla</b> for breakfast or an afternoon snack."
    ],
    recipe: {
      serves: "Makes 12", time: "30 min + about 2½ hr rising", tools: "Big bowl · baking sheet · butter knife",
      ingredients: ["Dough: 500 g (4 cups) all-purpose flour", "100 g (½ cup) sugar", "1 packet (2¼ tsp) instant yeast", "1 tsp salt", "2 eggs", "180 ml (¾ cup) warm milk", "115 g (½ cup) soft butter", "Topping: 65 g (½ cup) flour", "60 g (½ cup) powdered sugar", "55 g (¼ cup) soft butter", "1 tsp vanilla, or 1 tbsp cocoa for chocolate conchas"],
      steps: [
        ["Mix the flour, sugar, yeast and salt. Add the eggs, warm milk and butter, and knead until smooth and stretchy, about 10 minutes.", "together"],
        ["Cover the bowl and let the dough rise until doubled, about 1½ hours.", "kid"],
        ["Divide into 12 balls and set them on a lined baking sheet.", "kid"],
        ["Mix the topping ingredients into a soft paste. Flatten 12 little discs and press one onto each ball.", "kid"],
        ["Score shell lines into the topping with a butter knife. Let the rolls rise 45 minutes.", "kid"],
        ["Bake at 175 °C / 350 °F for 18–20 minutes, until puffed and lightly golden underneath.", "grown"]
      ]
    }
  },
  {
    week: 3, slug: "croissant", city: "Paris", country: "France",
    name: "Croissants", altName: "Croissant means “crescent”",
    tagline: "Flaky, buttery crescents",
    photoCredit: "",
    where: "France", whereSub: "Capital: Paris · Language: French",
    story: "The croissant’s ancestor is the Austrian <i>kipferl</i>, a crescent-shaped roll. In the late 1830s an Austrian baker named August Zang opened a Viennese bakery in Paris, and Parisians loved its crescents. French bakers later made them with layered, buttery dough, which became the flaky croissant we know today.",
    facts: [
      "<b>Croissant</b> means “crescent” in French, after its moon shape.",
      "Real croissant dough is folded again and again around butter to make dozens of thin layers. Bakers call this <b>laminating</b>.",
      "A proper bakery croissant takes about <b>two days</b> from start to finish."
    ],
    recipe: {
      serves: "Makes 8 (shortcut version)", time: "15 min + 18 min baking", tools: "Rolling pin · knife · baking sheet",
      note: "Real croissants take two days. This shortcut gets the flaky layers from ready-made puff pastry.",
      ingredients: ["1 sheet store-bought puff pastry, thawed", "8 small squares of chocolate or 2 tbsp jam (optional)", "1 egg, beaten", "1 tbsp sugar"],
      steps: [
        ["Unroll the pastry and cut it into 8 long triangles.", "together"],
        ["Put a little chocolate or jam at the wide end of each triangle.", "kid"],
        ["Roll each one up from the wide end to the point, then curve the ends into a crescent.", "kid"],
        ["Brush with beaten egg and sprinkle with sugar.", "kid"],
        ["Bake at 200 °C / 400 °F for 15–18 minutes until puffed and golden.", "grown"]
      ]
    }
  },
  {
    week: 4, slug: "melon-pan", city: "Tokyo", country: "Japan",
    name: "Melon Pan", altName: "メロンパン",
    tagline: "A sweet bun wearing a crunchy cookie coat",
    photoCredit: "",
    where: "Japan", whereSub: "Capital: Tokyo · Language: Japanese",
    story: "Melon pan is a soft, sweet bun covered in a crisp cookie crust. The crust is scored in a crisscross pattern that makes it look like a cantaloupe, but most melon pan doesn’t taste like melon at all! It has been a favorite at Japanese bakeries for about a century, and nobody agrees on exactly who invented it.",
    facts: [
      "It’s named for its <b>look, not its taste</b>. The crisscross top looks like a melon rind.",
      "You can find it in <b>bakeries and convenience stores</b> all over Japan.",
      "Some shops split it open and fill it with <b>ice cream</b>."
    ],
    recipe: {
      serves: "Makes 8", time: "40 min + 1½ hr rising", tools: "Bowls · baking sheet · butter knife",
      ingredients: ["Bun: 250 g (2 cups) bread flour", "3 tbsp sugar", "1½ tsp instant yeast", "½ tsp salt", "150 ml (⅔ cup) warm milk", "1 egg", "2 tbsp soft butter", "Cookie crust: 55 g (¼ cup) soft butter", "65 g (⅓ cup) sugar, plus extra for rolling", "½ beaten egg", "130 g (1 cup) flour", "½ tsp baking powder"],
      steps: [
        ["Mix and knead the bun ingredients until smooth, about 10 minutes. Let it rise 1 hour.", "together"],
        ["Beat the butter and sugar for the crust, add the egg, then mix in the flour and baking powder. Chill 20 minutes.", "kid"],
        ["Divide the bun dough into 8 balls. Divide the cookie dough into 8 and flatten each into a thin disc.", "kid"],
        ["Drape a cookie disc over each bun, roll the top in sugar, and score a crisscross pattern.", "kid"],
        ["Let rise 30–40 minutes, then bake at 180 °C / 350 °F for 13–15 minutes.", "grown"]
      ]
    }
  },
  {
    week: 5, slug: "pavlova", city: "Wellington", country: "New Zealand",
    name: "Pavlova", altName: "Named after a ballerina",
    tagline: "A cloud of meringue piled with cream and fruit",
    photoCredit: "",
    where: "New Zealand", whereSub: "Capital: Wellington · Languages: English, Māori and NZ Sign Language",
    story: "Pavlova is a big, cloud-like meringue that’s crisp outside and soft and marshmallowy inside, piled with whipped cream and fruit. It’s named after the Russian ballerina Anna Pavlova, who toured New Zealand and Australia in the 1920s.",
    facts: [
      "New Zealand and Australia have <b>argued for decades</b> about who made it first.",
      "In New Zealand it’s often topped with <b>kiwifruit</b>.",
      "It’s a favorite at <b>Christmas</b>, which is in summer there!"
    ],
    recipe: {
      serves: "Serves 8", time: "20 min + 3 hr baking and cooling", tools: "Electric mixer · baking sheet · parchment paper",
      ingredients: ["4 egg whites, at room temperature", "200 g (1 cup) superfine sugar", "1 tsp cornstarch", "1 tsp white vinegar", "1 tsp vanilla", "Topping: 240 ml (1 cup) cream, whipped", "2 kiwifruit and a handful of berries"],
      steps: [
        ["Heat the oven to 120 °C / 250 °F and line a baking sheet.", "grown"],
        ["Whip the egg whites until soft peaks form.", "together"],
        ["Add the sugar one spoonful at a time, whipping until thick and glossy.", "kid"],
        ["Gently fold in the cornstarch, vinegar and vanilla.", "kid"],
        ["Spoon it into a 20 cm (8 in) circle on the parchment, a little higher around the edge.", "kid"],
        ["Bake 1 hour 15 minutes, then turn off the oven and let it cool inside with the door closed.", "grown"],
        ["Just before serving, top with whipped cream and fruit.", "kid"]
      ]
    }
  },
  {
    week: 6, slug: "pastel-de-nata", city: "Lisbon", country: "Portugal",
    name: "Pastéis de Nata", altName: "One pastel, two pastéis",
    tagline: "Little custard tarts with caramelized tops",
    photoCredit: "",
    where: "Portugal", whereSub: "Capital: Lisbon · Language: Portuguese",
    story: "These little custard tarts were first made by monks at the Jerónimos Monastery in Belém, a neighborhood of Lisbon. When the monastery closed in 1834, the recipe was sold, and in 1837 a bakery nearby began selling them. It still does, and its recipe is kept secret.",
    facts: [
      "One story says monks used <b>egg whites to starch clothes</b>, leaving lots of yolks for sweets.",
      "People in Portugal sprinkle them with <b>cinnamon and powdered sugar</b>.",
      "The dark, spotty tops come from a <b>very hot oven</b>."
    ],
    recipe: {
      serves: "Makes 12", time: "30 min + 15 min baking", tools: "Muffin tin · saucepans · whisk",
      ingredients: ["1 sheet store-bought puff pastry", "240 ml (1 cup) whole milk", "3 tbsp flour", "130 g (⅔ cup) sugar", "80 ml (⅓ cup) water", "1 cinnamon stick", "1 strip of lemon peel", "5 egg yolks", "Cinnamon and powdered sugar for dusting"],
      steps: [
        ["Heat the oven as hot as it goes (about 245 °C / 475 °F) and grease a muffin tin.", "grown"],
        ["Roll the pastry into a tight log, cut it into 12 slices, and press each slice into a cup with your thumbs.", "kid"],
        ["Whisk the flour into a quarter of the milk. Heat the rest of the milk, then whisk the two together until thick.", "together"],
        ["Boil the sugar, water, cinnamon and lemon peel for 3 minutes to make a syrup.", "grown"],
        ["Whisk the syrup into the milk mix and let it cool a little, then whisk in the yolks.", "together"],
        ["Fill each pastry cup three-quarters full.", "kid"],
        ["Bake 12–15 minutes until the tops have dark spots. Dust with cinnamon.", "grown"]
      ]
    }
  },
  {
    week: 7, slug: "sachertorte", city: "Vienna", country: "Austria",
    name: "Sachertorte", altName: "Invented by a 16-year-old",
    tagline: "Chocolate cake, apricot jam, shiny chocolate glaze",
    photoCredit: "",
    where: "Austria", whereSub: "Capital: Vienna · Language: German",
    story: "In 1832, Prince Metternich asked his kitchen for a special dessert. The head chef was sick, so 16-year-old apprentice Franz Sacher invented a chocolate cake with apricot jam under a shiny chocolate glaze. His family later ran Vienna’s Hotel Sacher, where it’s still served.",
    facts: [
      "The inventor, <b>Franz Sacher, was only 16</b>.",
      "Hotel Sacher and a famous Vienna bakery, Demel, once went to <b>court</b> over who could call theirs the “Original Sachertorte.”",
      "In Vienna it’s served with <b>unsweetened whipped cream</b>."
    ],
    recipe: {
      serves: "Serves 10", time: "45 min + 1 hr baking and cooling", tools: "20 cm (8 in) cake pan · mixer · saucepan",
      ingredients: ["115 g (½ cup) soft butter", "100 g (½ cup) sugar, divided", "4 eggs, separated", "115 g (4 oz) dark chocolate, melted", "95 g (¾ cup) flour", "Pinch of salt", "160 g (½ cup) apricot jam, warmed", "Glaze: 115 g (4 oz) dark chocolate and 120 ml (½ cup) cream"],
      steps: [
        ["Heat the oven to 175 °C / 350 °F and grease the pan.", "grown"],
        ["Beat the butter with half the sugar, then beat in the yolks and melted chocolate.", "together"],
        ["Whip the egg whites with the rest of the sugar until they hold soft peaks.", "together"],
        ["Gently fold the whites and the flour into the chocolate mix.", "kid"],
        ["Bake 40–45 minutes, then let it cool completely.", "grown"],
        ["Spread the warm apricot jam over the top.", "kid"],
        ["Heat the cream, stir in the chocolate until smooth, and pour it over the cake.", "together"]
      ]
    }
  },
  {
    week: 8, slug: "cannoli", city: "Palermo", country: "Italy",
    name: "Cannoli", altName: "One cannolo, two cannoli",
    tagline: "Crunchy pastry tubes with sweet ricotta cream",
    photoCredit: "",
    where: "Sicily, the island at the tip of Italy’s boot", whereSub: "Sicily’s capital: Palermo · Languages: Italian and Sicilian",
    story: "Cannoli are crunchy fried pastry tubes filled with sweet, creamy ricotta. They come from Sicily, where they were made as a special treat for Carnevale, the festival before Lent.",
    facts: [
      "<b>Cannoli</b> is the plural. Just one is a <b>cannolo</b>.",
      "Traditional Sicilian ricotta is made from <b>sheep’s milk</b>.",
      "Fill them <b>right before eating</b> so the shells stay crunchy."
    ],
    recipe: {
      serves: "Makes 12", time: "20 min + draining the ricotta", tools: "Strainer · bowl · piping bag or zip-top bag",
      note: "Frying the shells is hot, tricky work, so this version uses store-bought shells.",
      ingredients: ["500 g (2 cups) whole-milk ricotta", "90 g (¾ cup) powdered sugar, plus extra for dusting", "1 tsp vanilla", "Zest of 1 orange", "½ cup mini chocolate chips", "12 store-bought cannoli shells", "Chopped pistachios (optional)"],
      steps: [
        ["Drain the ricotta in a strainer for at least an hour so the filling is thick.", "together"],
        ["Mix the ricotta, powdered sugar, vanilla and orange zest until smooth.", "kid"],
        ["Stir in the chocolate chips.", "kid"],
        ["Spoon the filling into a piping bag or zip-top bag and snip off a corner.", "together"],
        ["Fill each shell from both ends, dip the ends in pistachios, and dust with powdered sugar.", "kid"]
      ]
    }
  },
  {
    week: 9, slug: "kanelbulle", city: "Stockholm", country: "Sweden",
    name: "Kanelbullar", altName: "Swedish cinnamon buns",
    tagline: "Cardamom-scented cinnamon buns with pearl sugar",
    photoCredit: "",
    where: "Sweden", whereSub: "Capital: Stockholm · Language: Swedish",
    story: "Kanelbullar are a big part of <i>fika</i>, the Swedish habit of pausing for a drink and something sweet with friends. Swedish buns are less gooey than American cinnamon rolls and are often flavored with cardamom and topped with crunchy pearl sugar.",
    facts: [
      "<b>October 4</b> is <i>Kanelbullens dag</i>, Cinnamon Bun Day, in Sweden.",
      "<b>Fika</b> is so important that many Swedish workplaces have a fika break.",
      "<b>Pearl sugar</b> is little white nuggets of sugar that don’t melt in the oven."
    ],
    recipe: {
      serves: "Makes 16", time: "40 min + 1½ hr rising", tools: "Rolling pin · knife · muffin tin or paper cups",
      ingredients: ["Dough: 240 ml (1 cup) warm milk", "1 packet (2¼ tsp) instant yeast", "65 g (⅓ cup) sugar", "1 tsp ground cardamom", "½ tsp salt", "85 g (6 tbsp) soft butter", "440 g (3½ cups) flour", "Filling: 115 g (½ cup) soft butter", "100 g (½ cup) brown sugar", "1 tbsp cinnamon", "1 egg, beaten, and pearl sugar (or crushed sugar cubes)"],
      steps: [
        ["Mix the dough ingredients and knead about 8 minutes. Let it rise 1 hour.", "together"],
        ["Roll the dough into a big rectangle, about 30 × 40 cm (12 × 16 in).", "together"],
        ["Spread the filling all over the dough.", "kid"],
        ["Roll it up tightly from the long side and cut into 16 slices.", "together"],
        ["Set each slice in a paper cup or muffin tin and let rise 30 minutes.", "kid"],
        ["Brush with egg and sprinkle with pearl sugar.", "kid"],
        ["Bake at 220 °C / 425 °F for 8–10 minutes until golden.", "grown"]
      ]
    }
  },
  {
    week: 10, slug: "baklava", city: "Istanbul", country: "Türkiye",
    name: "Baklava", altName: "Layers, nuts and syrup",
    tagline: "Paper-thin pastry, chopped nuts and sweet syrup",
    photoCredit: "",
    where: "Türkiye", whereSub: "Capital: Ankara · Largest city: Istanbul · Language: Turkish",
    story: "Baklava is made from paper-thin layers of pastry, chopped nuts and butter, soaked in syrup after baking. It was perfected in the kitchens of the Ottoman sultans at Topkapı Palace in Istanbul.",
    facts: [
      "During Ramadan, the sultan’s kitchen sent trays of baklava to the palace soldiers in a parade called the <b>Baklava Procession</b>.",
      "The city of <b>Gaziantep</b> in southern Türkiye is famous for its pistachio baklava.",
      "Many countries, from <b>Greece to the Middle East</b>, have their own version."
    ],
    recipe: {
      serves: "Makes about 24 small pieces", time: "45 min + 45 min baking + 4 hr resting", tools: "20 cm (8 in) square pan · pastry brush · saucepan",
      ingredients: ["1 package phyllo dough, thawed", "115 g (½ cup) butter, melted", "180 g (1½ cups) walnuts or pistachios, finely chopped", "1 tsp cinnamon", "Syrup: 150 g (¾ cup) sugar", "120 ml (½ cup) water", "1 tbsp honey", "1 tsp lemon juice"],
      steps: [
        ["Simmer the syrup ingredients for 10 minutes, then let it cool.", "grown"],
        ["Mix the nuts and cinnamon.", "kid"],
        ["Heat the oven to 175 °C / 350 °F. Keep the phyllo covered with a damp towel so it doesn’t dry out.", "grown"],
        ["Layer 8 sheets in the pan, brushing each with butter. Sprinkle half the nuts, add 4 more buttered sheets, then the rest of the nuts.", "together"],
        ["Finish with 8 more buttered sheets on top.", "kid"],
        ["Cut into diamonds before baking, then bake 40–45 minutes until golden.", "grown"],
        ["Pour the cool syrup over the hot baklava. Then comes the hardest part: waiting 4 hours!", "kid"]
      ]
    }
  },
  {
    week: 11, slug: "alfajores", city: "Buenos Aires", country: "Argentina",
    name: "Alfajores", altName: "Alfajores de maicena",
    tagline: "Melt-in-your-mouth cookies sandwiching dulce de leche",
    photoCredit: "",
    where: "Argentina", whereSub: "Capital: Buenos Aires · Language: Spanish",
    story: "Alfajores are two soft, crumbly cookies sandwiching <i>dulce de leche</i>, a thick caramel made by slowly cooking milk and sugar. The word came into Spanish from Arabic, and the treat traveled to South America with Spanish settlers. Today they’re one of Argentina’s favorite snacks.",
    facts: [
      "<b>Maicena</b> means cornstarch. It makes the cookies melt in your mouth.",
      "The edges are often rolled in <b>shredded coconut</b>.",
      "<b>Dulce de leche</b> means “sweet made from milk.”"
    ],
    recipe: {
      serves: "Makes about 15 sandwiches", time: "40 min + 30 min chilling", tools: "Mixer · rolling pin · round cutter",
      ingredients: ["130 g (1 cup) flour", "150 g (1¼ cups) cornstarch", "1 tsp baking powder", "115 g (½ cup) soft butter", "65 g (⅓ cup) sugar", "2 egg yolks", "1 tsp vanilla", "Zest of ½ lemon", "1 cup dulce de leche", "½ cup shredded coconut"],
      steps: [
        ["Beat the butter and sugar, then add the yolks, vanilla and lemon zest.", "kid"],
        ["Mix in the flour, cornstarch and baking powder. Chill the dough 30 minutes.", "kid"],
        ["Roll it out about 6 mm (¼ in) thick and cut small circles.", "kid"],
        ["Bake at 175 °C / 350 °F for 10–12 minutes, until set but still pale.", "grown"],
        ["When cool, sandwich two cookies with a spoonful of dulce de leche.", "kid"],
        ["Roll the edges in coconut.", "kid"]
      ]
    }
  },
  {
    week: 12, slug: "skolebrod", city: "Oslo", country: "Norway",
    name: "Skolebrød", altName: "“School bread”",
    tagline: "Cardamom buns with a custard center and coconut",
    photoCredit: "",
    where: "Norway", whereSub: "Capital: Oslo · Language: Norwegian",
    story: "<i>Skolebrød</i> means “school bread.” It’s a soft cardamom bun with a pocket of vanilla custard in the middle, topped with icing and coconut. It became a favorite in Norwegian school lunches and bakeries.",
    facts: [
      "Norwegians also call them <b>skoleboller</b>, “school buns.”",
      "<b>Cardamom</b> is one of Scandinavia’s favorite baking spices.",
      "The custard stays put because you press a <b>deep dent</b> into each bun before baking."
    ],
    recipe: {
      serves: "Makes 12", time: "1 hr + 1½ hr rising", tools: "Saucepan · whisk · baking sheet",
      ingredients: ["Dough: 240 ml (1 cup) warm milk", "1 packet (2¼ tsp) instant yeast", "65 g (⅓ cup) sugar", "1 tsp ground cardamom", "½ tsp salt", "85 g (6 tbsp) soft butter", "440 g (3½ cups) flour", "Custard: 240 ml (1 cup) milk, 3 egg yolks, 50 g (¼ cup) sugar, 2 tbsp cornstarch, 1 tsp vanilla", "Icing: 120 g (1 cup) powdered sugar and 1–2 tbsp water", "½ cup shredded coconut"],
      steps: [
        ["Cook the custard ingredients over low heat, whisking, until thick. Chill it.", "together"],
        ["Mix and knead the dough, then let it rise 1 hour.", "together"],
        ["Shape 12 balls and let them rise 30 minutes.", "kid"],
        ["Press a deep dent into each bun and fill it with custard.", "kid"],
        ["Bake at 220 °C / 425 °F for 10–12 minutes.", "grown"],
        ["When cool, spread icing around the custard and dip in coconut.", "kid"]
      ]
    }
  }
];
