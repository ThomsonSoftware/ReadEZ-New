Option Strict On

Imports System.IO

' ====================================================================================================
'   ReadingTexts.vb - the texts read in the symptoms and rate-of-reading tests
'
'   The built-in texts are those of ReadEZ 2020, in the same order, so the text number saved with
'   a result (RoR_TextOptions) still names the same text:
'     0       Wilkins Rate of Reading (random high-frequency words, as designed by Prof Wilkins)
'     1 - 4   random words (levels 1 and 2), numbers and letters
'     5 - 10  graded passages, 11 - 12 prose, 13 German, 14 Afrikaans
'   Extra texts can be added as *.txt files (one row per line) in the data folder's
'   "Reading texts" folder (Settings > Test setup); they follow the built-in ones in the list.
' ====================================================================================================

Public NotInheritable Class ReadingText
    Public ReadOnly Property Name As String
    Public ReadOnly Property Rows As IReadOnlyList(Of String)
    ''' <summary>True for a text from the data folder (it can be edited or deleted).</summary>
    Public ReadOnly Property IsUserText As Boolean

    Public Sub New(name As String, rows As String(), Optional isUserText As Boolean = False)
        Me.Name = name
        Me.Rows = rows
        Me.IsUserText = isUserText
    End Sub

    Public Overrides Function ToString() As String
        Return Name
    End Function
End Class

Public NotInheritable Class ReadingTexts
    Private Sub New()
    End Sub

    Private Shared ReadOnly BuiltIn As ReadingText() = {
        New ReadingText("Wilkins Rate of Reading", {
            "come see the play look up is cat not my and dog for you to",
            "the cat up dog and is play come you see for not to look my",
            "you for the and not see my play come is look dog cat to up",
            "dog to you and play cat up is my not come for the look see",
            "play come see cat not look dog is my up the for to and you",
            "to not cat for look is my and up come play you see the dog",
            "my play see to for you is the look up cat not dog come and",
            "look to for my come play the dog see you not cat up and is",
            "up come look for the not dog cat you to see is and my play",
            "is you dog for not cat my look come and up to play see the",
            "see the look dog and not is you come up to my for cat play",
            "not up play my is dog you come look for see and to the cat",
            "look up come and is my cat not dog you see for to play the",
            "my you is look the dog play see not come and to cat for up",
            "for the to and you cat is look up my not dog play see come",
            "you look see and play to the is cat not come for my up dog",
            "come not to play look the and dog see is cat up you for my",
            "and is for dog come see the cat up look you play my not to",
            "dog you cat to and play for not come up the see look my is",
            "the come to up cat my see dog you not look is play and for"}),
        New ReadingText("Random Words Level 1", {
            "they one had some word go or all not my and out way but in",
            "had all go out and or some they but one way not in word my",
            "but way had and not one my some they or word out all in no",
            "out in but had some all go or my not they way and word one",
            "some they one all not word out or my go had way in and but",
            "in not all way word or my had go they some but one and out",
            "my some one in way but or had word go all not out they and",
            "word in way my they some had out one but not all go and or",
            "go they word way had not out all but in one or and my some",
            "or but out way not all my word they had go in some one the",
            "one and word out had not or but they go in my way all some",
            "not go some my or out but they word way one and in had all",
            "word go they and or my all not out but one way in some the",
            "my but or word had out some one not they and in all way no",
            "way had in and but all or word go my not out some one they",
            "but word one had some in and or all not they way my go out",
            "they not in some word had and out one or all go but way my",
            "had or way out they one and all go word but some my not in",
            "out but all in had some way not they go and one word my or",
            "had they in go all my one out but not word or some and way"}),
        New ReadingText("Random Words Level 2", {
            "line put big kind does say why ask try still same high much well any",
            "big ask say high same why kind line well put much try any does still",
            "well much big same try put still kind line why does high ask any say",
            "high any well same kind ask say why still try line much big does put",
            "kind line put ask try does high why still say big much any same well",
            "any try ask much does why still same say line kind well put big high",
            "still kind put any much well why big does say ask try high line same",
            "does any much still line kind big high put well try ask say same why",
            "say line does much big try high ask well any put why same still kind",
            "why well high much try ask still does line same say any kind put big",
            "put big does high same try why well line say any still much ask kind",
            "try say kind still why high well line does much put same any big ask",
            "does say line same why still ask try high well put much any kind big",
            "still well why does big high kind put try line same any ask much say",
            "much big any same well ask why does say still try high kind put line",
            "well does put same kind any big why ask try line much still say high",
            "line try any kind does big same high put why ask say well much still",
            "same why much high line put big ask say does well kind still try any",
            "high well ask any same kind much try line say big put does still why",
            "big line any say ask still put high well try does why kind same much"}),
        New ReadingText("Random Numbers", {
            "24 12 39 43 91 55 68 74 87 31 79 54 28 46 10",
            "39 74 55 54 79 68 43 24 46 12 28 87 10 91 31",
            "46 28 39 79 87 12 31 43 24 68 91 54 74 10 55",
            "54 10 46 79 43 74 55 68 31 87 24 28 39 91 12",
            "43 24 12 74 87 91 54 68 31 55 39 28 10 79 46",
            "10 87 74 28 91 68 31 79 55 24 43 46 12 39 54",
            "31 43 12 10 28 46 68 39 91 55 74 87 54 24 79",
            "91 10 28 31 24 43 39 54 12 46 87 74 55 79 68",
            "55 24 91 28 39 87 54 74 46 10 12 68 79 31 43",
            "68 46 54 28 87 74 31 91 24 79 55 10 43 12 39",
            "12 39 91 54 79 87 68 46 24 55 10 31 28 74 43",
            "87 55 43 31 68 54 46 24 91 28 12 79 10 39 74",
            "91 55 24 79 68 31 74 87 54 46 12 28 10 43 39",
            "31 46 68 91 39 54 43 12 87 24 79 10 74 28 55",
            "28 39 10 79 46 74 68 91 55 31 87 54 43 12 24",
            "46 91 12 79 43 10 39 68 74 87 24 28 31 55 54",
            "24 87 10 43 91 39 79 54 12 68 74 55 46 28 31",
            "79 68 28 54 24 12 39 74 55 91 46 43 31 87 10",
            "54 46 74 10 79 43 28 87 24 55 39 12 91 31 68",
            "39 24 10 55 74 31 12 54 46 87 91 68 43 79 28"}),
        New ReadingText("Random Words Letters", {
            "c a e b i n o z r s u v x d h",
            "e z n v u o b c d a x r h i s",
            "d x e u r a s b c o i v z h n",
            "v h d u b z n o s r c x e i a",
            "b c a z r i v o s n e x h u d",
            "h r z x i o s u n c b d a e v",
            "s b a h x d o e i n z r v c u",
            "i h x s c b e v a d r z n u o",
            "n c i x e r v z d h a o u s b",
            "o d v x r z s i c u n h b a e",
            "a e i v u r o d c n h s x z b",
            "r n b s o v d c i x a u h e z",
            "i n c u o s z r v d a x h b e",
            "s d o i e v b a r c u h z x n",
            "x e h u d z o i n s r v b a c",
            "d i a u b h e o z r c x s n v",
            "c r h b i e u v a o z n d x s",
            "u o x v c a e z n i d b s r h",
            "v d z h u b x r c n e a i s o",
            "e c h n z s a v d r i o b u x"}),
        New ReadingText("Level 1", {
            "Bob looked down the street. A man was riding a gray pony.",
            "Five pence a ride, said the man. Bob got five pence from his mother.",
            "He went for a ride down the street. Then Bob came back on the pony.",
            "Stay on the pony, said the man. He took Bob’s picture on the pony.",
            "Bob gave his mother the picture. Jane had a new white kitten.",
            "She took the kitten out to play. Then two big black dogs ran by.",
            "The little kitten was afraid of dogs. Up a tree went the little kitten.",
            "Jane called and called to her kitten. The kitten did not come down to Jane.",
            "Mother put some milk under the tree. The kitten came down for the milk.",
            "One day Judy opened her window. A green bird was at the window.",
            "He hopped in and sat on Judy’s hand. Mother said, He must be a pet.",
            "Judy said, We must find his home. Four children came to look at him.",
            "But he was not their green bird. Judy said, Can the bird stay here?",
            "Mother said, Yes, we will get a cage. Mary was playing in the yard.",
            "She was looking for pretty red leaves. She saw something brown in the leaves.",
            "She thought it was a brown coat. Then it began to hop and jump around",
            "Oh, said Mary. It is a rabbit. Mary ran to catch the little rabbit.",
            "But the little rabbit ran too fast. It hopped into a hole."}),
        New ReadingText("Level 2", {
            "Dick was visiting the fire station. Suddenly the fire bell gave three rings.",
            "All the firemen jumped on the fire engine and drove off.  Dick stayed at",
            "the fire station. Soon he began to smell something funny. He saw a small fire",
            "under the table. Dick ran for a bucket of water and put the fire out.",
            "The firemen were proud of him. Sally liked to write letters.",
            "She would write a long letter to one friend each day. One day, she heard",
            "the doorbell ring. The postman gave Sally a letter. She said to him,",
            "this letter must be from my friend Ann. But the postman shook his head.",
            "He said, this is your own letter. You forgot to write your friend’s name.",
            "Alan was watching television. The show was called Stop the Bell.",
            "You had to answer a question before the bell started to ring. You got a",
            "prize if you were right. Alan was surprised to see his teacher on the",
            "show. She answered all of her questions right. Her prize was a yellow car.",
            "All of the children were proud of her. Peggy took a bus trip to the city.",
            "She was taking a present to her aunt. Peggy hurried off the bus to",
            "meet her aunt. As they left the station, Peggy thought of the present. It",
            "was still on the bus. Just then a man came running up with the present.",
            "You left this behind, said the man. Oh, thank you, said Peggy."}),
        New ReadingText("Level 3", {
            "David likes to save stamps. He has saved over two thousand stamps.",
            "Sometimes he buys them, but he takes most of them off letters.  David has",
            "a special book for his stamps with one page for each county.  One of his",
            "stamps is worth fifteen dollars. This one shows an airplane that is upside down.",
            "David likes to trade stamps. Linda likes to save postcards.",
            "She has friends in many countries who send her picture postcards. So far, she",
            "has cards from fifteen countries. All the cards are in colour, and the pictures",
            "show people and buildings. Linda likes the pictures of castles best. When she grows",
            "up she plans to travel around the world. She wants a postcard from every country.",
            "Ellen likes to save playing cards. She saves only cards with animal pictures.",
            "Her friend Jane saves cards with birds and flowers. Ellen likes to trade cards",
            "with her friends. The cards that are worth the most are those with many",
            "bright colours. The card Ellen likes the best shows a picture of a funny brown camel.",
            "She has over fifty different cards. Jean has over a hundred dolls.",
            "When her father returns from a trip, he brings Jean another doll. She handles",
            "her dolls very carefully and keeps them in a special box in her bedroom.",
            "She has dolls from many different lands. One doll wears a purple ribbon in her",
            "hair. She wears gloves on her hands. Jean calls this doll Princess Pat."}),
        New ReadingText("Level 5", {
            "Whales breathe air just as human beings do. But they do not breathe",
            "through their mouths as we can.  Instead they breathe through",
            "two little nostrils placed on top of their heads.  When the whale",
            "sinks under the water, special muscles shut his nose to keep out",
            "the water.  This makes it possible for him to swim with his mouth",
            "open.  He can capture prey and even eat under water without danger",
            "of drowning.  A whale can blow the hot air from his lungs with",
            "tremendous force. People used to think this fountain of steam was",
            "water. His eyes are set high for seeing at far distances. They are",
            "protected by overhanging lids and by long lashes which cover them",
            "from the blinding sun. His nose can be closed tight against the",
            "driving sand of the desert storms. The camel’s upper lip reaches",
            "out for shrubs and other plants.  The camel is not particular about",
            "its food and its strong yellow teeth can chew just about anything,",
            "even bitter, thorny shrubs.  Many of them will even eat leather,",
            "blankets or bones and will also drink salty water. Nothing seems",
            "to give a camel indigestion. The daddy longlegs is a curious looking",
            "insect. Though he resembles a spider, he is really a closer relative"}),
        New ReadingText("Level 6", {
            "Each year an exciting contest is held in Alaska. Everyone attempts",
            "to guess the exact minute and day the ice will break in one of the",
            "important rivers. This happens in late April or early May. The",
            "contest starts after January first each year. Anyone may pay a",
            "dollar and have their guess posted in the record book. To find out the",
            "precise minute the ice breaks, wires are stretched to the middle of",
            "the river and frozen through the ice. These wires are attached to a",
            "timing device and bell on the shore. Crowds of people gather to",
            "listen for the bell, announcing the cracking of the ice. The closest",
            "guess wins thousands of dollars. Brazil grows over fifty per cent of",
            "the world's coffee beans. Each company’s brand of coffee is",
            "blended from several kinds of coffee beans. As the beans may vary in",
            "taste, each company has a coffee taster whose job it is to make",
            "sure that his company’s brand always tastes the same. He sits at",
            "a revolving table filled with numerous cups of black coffee. He takes",
            "a sip from each cup, tastes carefully, then spits it out. To keep his",
            "taste working properly, he must never drink the coffee. The coffee",
            "beans are shipped to the companies all around the world in big ships"}),
        New ReadingText("Level 10", {
            "Amundsen of Norway was one of the great explorers. He was noted for his",
            "thoroughness and perseverance. On his first expedition, he located the",
            "north magnetic pole and found the Northwest Passage. On his next voyage,",
            "he headed for the North Pole but, hearing that another explorer had",
            "reached it, sailed for the South Pole instead. On arriving, Amundsen",
            "established his base station on the ice cap. Using dog teams and sledges",
            "he and his men equipped three supply stations to the south. After camping",
            "for the winter, they started for the pole, a journey of two months’ duration.",
            "In 1911, Amundsen flew the Norwegian flag above the pole. They spent a total",
            "of eleven months reaching the South Pole. Born Ehrich Weiss in 1874, he took",
            "his professional name from the French magician, Houdini. No lock or straight",
            "jacket, no handcuff or sealed chest could restrain Harry Houdini. To secure",
            "publicity, he would vanish from the prison of each city where he appeared.",
            "A modest man, he performed his tricks with apparent ease, stating that they",
            "were based on natural laws. He was angered by mediums and others who claimed",
            "to have supernatural power, and spent much time and energy in exposing such",
            "frauds. An authority on magic, he willed his collection of writing to the",
            "Library of Congress. It is one of the most complete in the world. Louis"}),
        New ReadingText("The pet shop", {
            "Cody and his sister April decide they want a dog. They head",
            "down to the local pet store and have a look around. It is a",
            "very small pet store that doesn't have many animals. The",
            "owner of the shop is a nice old man named Mr. Smith. He walks",
            "over and greets Cody and April. How can I help you? he asks.",
            "We would like to buy a dog, April responds. Ah, well, we are",
            "not a big pet shop, Mr. Smith tells her. So we only have two",
            "dogs to choose from. They ask Mr. Smith to show them the dogs.",
            "Mr. Smith leads them to the back of store where the two dogs",
            "are. One of them is a very big bulldog named Buster. The other",
            "is a very tiny poodle named Teacup. April wants Teacup. Cody",
            "wants Buster. They walk outside to discuss.They can’t agree",
            "on a dog. April suggests they race home for it. The winner",
            "of the race chooses the dog. Cody agrees, then tells April",
            "her shoelace is untied. When April looks down, he runs off",
            "and gets a head start. Cody runs as hard as he can. He really",
            "wants that bulldog. He looks back. April is so far behind he",
            "can’t even see her. Cody finally gets home. He is tired but",
            "he is happy. He knows he is the winner.April arrives a few",
            "minutes after Cody. She congratulates him. They return to the"}),
        New ReadingText("The moon", {
            "The Moon is an astronomical body that orbits planet Earth,",
            "being Earth's only permanent natural satellite. It is the",
            "fifth-largest natural satellite in the Solar System, and",
            "the largest among planetary satellites relative to the size",
            "of the planet that it orbits (its primary). Following",
            "Jupiter's satellite Io, the Moon is second-densest satellite",
            "among those whose densities are known.The average distance",
            "of the Moon from the Earth is 384,400 km (238,900 mi), or",
            "1.28 light-seconds.The Moon is thought to have formed about",
            "4 billion years ago, not long after Earth. There are several",
            "hypotheses for its origin; the most widely accepted explanation",
            "is that the Moon formed from the debris left over after a",
            "giant impact between Earth and a Mars-sized body called Theia.",
            "The Moon is in synchronous rotation with Earth, always showing",
            "the same face, with its near side marked by dark volcanic maria",
            "that fill the spaces between the bright ancient crustal",
            "highlands and the prominent impact craters. It is the second-",
            "brightest regularly visible celestial object in Earth's sky,",
            "after the Sun,, as measured by illuminance on Earth's surface.",
            "Its surface is actually dark, although compared to the night"}),
        New ReadingText("German", {
            "ja auf und ist vom lesen das bei nur alle du uns los mit ein",
            "los vom ist auf ja lesen alle mit ein und bei das nur du uns",
            "vom nur alle mit ist und bei uns das lesen ein du los ja auf",
            "alle uns ein und ja ist vom du los mit auf lesen bei das nur",
            "ein vom los mit uns du alle nur bei das lesen ist und auf ja",
            "uns mit alle bei lesen ist auf ja und vom das nur du ein los",
            "und lesen nur mit los uns du bei vom auf ja ist das ein alle",
            "mit nur lesen und ja ist das alle ein los du bei vom uns auf",
            "nur das vom auf bei alle ein lesen mit du uns los ist ja und",
            "lesen bei alle ein ja ist auf nur und mit uns das los vom du",
            "los auf ja lesen alle mit ein und bei das nur du uns vom ist",
            "bei lesen uns mit alle ist auf ja und vom das nur du ein los",
            "bei alle ein ja ist auf nur und mit uns lesen das los vom du",
            "ist das ein alle und lesen nur mit los uns du bei vom auf ja",
            "nur und mit lesen bei alle ein ja ist auf uns das los vom du",
            "ist auf los vom ja lesen alle mit ein und bei das nur du uns",
            "du ja auf und ist vom lesen das bei nur alle uns los mit ein",
            "mit nur lesen und ja ist das alle ein los du bei vom uns auf",
            "nur das vom auf bei alle ein lesen mit du uns los ist ja und",
            "lesen bei alle ein ja ist auf nur und mit uns das los vom du"}),
        New ReadingText("Afrikaans", {
            "boom nee vat slot keel op is kat nie my vir dak sit jou en",
            "vat kat op dak vir is slot boom jou nee sit nie en keel my",
            "jou sit vat vir nie nee my slot boom is keel dak kat en op",
            "dak en jou vir slot kat op is my nie boom sit vat keel nee",
            "slot boom nee kat nie keel dak is my op vat sit en vir jou",
            "en nie kat sit keel is my vir op boom slot jou nee vat dak",
            "my slot nee en sit jou is vat keel op kat nie dak boom vir",
            "keel en sit my boom slot vat dak nee jou nie kat op vir is",
            "op boom keel sit vat nie dak kat jou en nee is vir my slot",
            "is jou dak sit nie kat my keel boom vir op en slot nee vat",
            "nee vat keel dak vir nie is jou boom op en my sit kat slot",
            "nie op slot my is dak jou boom keel sit nee vir en vat kat",
            "keel op boom vir is my kat nie dak jou nee sit en slot vat",
            "my jou is keel vat dak slot nee nie boom vir en kat sit op",
            "sit vat en vir jou kat is keel op my nie dak slot nee boom",
            "jou keel nee vir slot en vat is kat nie boom sit my op dak",
            "boom nie en slot keel vat vir dak nee is kat op jou sit my",
            "vir is sit dak boom nee vat kat op keel jou slot my nie en",
            "dak jou kat en vir slot sit nie boom op vat nee keel my is",
            "vat boom en op kat my nee dak jou nie keel is slot vir sit"})}

    ''' <summary>Number of texts that come with ReadEZ.</summary>
    Public Shared ReadOnly Property BuiltInCount As Integer
        Get
            Return BuiltIn.Length
        End Get
    End Property

    ''' <summary>The name of a built-in text ("" if the number isn't one).</summary>
    Public Shared Function BuiltInName(index As Integer) As String
        Return If(index >= 0 AndAlso index < BuiltIn.Length, BuiltIn(index).Name, "")
    End Function

    ''' <summary>The folder for extra texts (data folder\Reading texts).</summary>
    Public Shared ReadOnly Property UserFolder As String
        Get
            Return Path.Combine(AppHost.Settings.DataDirectory, "Reading texts")
        End Get
    End Property

    ''' <summary>The built-in texts followed by the data folder's own (sorted by name).</summary>
    Public Shared Function All() As List(Of ReadingText)
        Dim list As New List(Of ReadingText)(BuiltIn)
        Try
            If Directory.Exists(UserFolder) Then
                For Each f In Directory.GetFiles(UserFolder, "*.txt").OrderBy(Function(x) x, StringComparer.CurrentCultureIgnoreCase)
                    Dim rows = File.ReadAllLines(f).Select(Function(l) l.Trim()).Where(Function(l) l <> "").ToArray()
                    If rows.Length > 0 Then list.Add(New ReadingText(Path.GetFileNameWithoutExtension(f), rows, isUserText:=True))
                Next
            End If
        Catch ex As Exception
            Log.Error("Reading texts", ex)
        End Try
        Return list
    End Function

    ''' <summary>The text with this number (the first one if it no longer exists).</summary>
    Public Shared Function ByIndex(index As Integer) As ReadingText
        Dim list = All()
        Return If(index >= 0 AndAlso index < list.Count, list(index), list(0))
    End Function

    ''' <summary>
    ''' The rows shown: the first n rows in order or, in mixed order (ReadEZ 2020's "random"), rows
    ''' taken in a fixed shuffled order so a retest shows the same layout.
    ''' </summary>
    Public Shared Function RowsFor(text As ReadingText, count As Integer, mixed As Boolean) As List(Of String)
        Dim order = Enumerable.Range(0, text.Rows.Count).ToList()
        If mixed Then
            Dim rnd As New Random(1234)
            For i = order.Count - 1 To 1 Step -1
                Dim j = rnd.Next(i + 1)
                Dim t = order(i) : order(i) = order(j) : order(j) = t
            Next
        End If
        Return order.Take(Math.Max(1, Math.Min(count, order.Count))).Select(Function(i) text.Rows(i)).ToList()
    End Function
End Class
